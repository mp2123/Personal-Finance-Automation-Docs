import os
import datetime
import logging
import plaid
from plaid.api import plaid_api
from plaid.model.accounts_get_request import AccountsGetRequest
from plaid.model.country_code import CountryCode
from plaid.model.item_public_token_exchange_request import ItemPublicTokenExchangeRequest
from plaid.model.institutions_get_by_id_request import InstitutionsGetByIdRequest
from plaid.model.item_get_request import ItemGetRequest
from plaid.model.link_token_create_request import LinkTokenCreateRequest
from plaid.model.link_token_create_request_user import LinkTokenCreateRequestUser
from plaid.model.products import Products
from plaid.model.transactions_get_request import TransactionsGetRequest
from plaid.model.transactions_get_request_options import TransactionsGetRequestOptions
from plaid.exceptions import ApiException as PlaidApiException

from .account_labels import build_account_label

logger = logging.getLogger(__name__)

class PlaidClient:
    def __init__(self, client_id, secret, env='sandbox'):
        self.client_id = client_id
        self.secret = secret
        self.env = env
        
        # Map environment string to Plaid Environment object
        host = plaid.Environment.Sandbox
        if self.env == 'development' and hasattr(plaid.Environment, 'Development'):
            host = plaid.Environment.Development
        elif self.env == 'production':
            host = plaid.Environment.Production
        
        self.config = plaid.Configuration(
            host=host,
            api_key={
                'clientId': self.client_id,
                'secret': self.secret,
            }
        )
        self.api_client = plaid.ApiClient(self.config)
        self.client = plaid_api.PlaidApi(self.api_client)

    def get_transactions(self, access_token, start_date, end_date):
        """Fetches transactions from Plaid for the specified date range."""
        transactions = []
        try:
            request = TransactionsGetRequest(
                access_token=access_token,
                start_date=start_date,
                end_date=end_date,
                options=TransactionsGetRequestOptions(
                    count=500, # Max allowed per request
                    offset=0,
                    include_personal_finance_category=True
                )
            )
            # Some versions of the SDK might handle this differently, 
            # but we explicitly request PFCs.
            # If the SDK supports it as a direct parameter:
            # request.personal_finance_category_version = 'v2'
            response = self.client.transactions_get(request)
            transactions.extend(response['transactions'])

            # Handle pagination if necessary
            total_transactions = response['total_transactions']
            while len(transactions) < total_transactions:
                logger.info(f"Fetching more transactions, offset: {len(transactions)}")
                request.options.offset = len(transactions)
                response = self.client.transactions_get(request)
                transactions.extend(response['transactions'])
                if not response['transactions']: 
                     logger.warning("Received empty transaction list during pagination, stopping fetch.")
                     break

            logger.info(f"Successfully retrieved {len(transactions)} transactions.")
            return transactions

        except PlaidApiException as e:
            logger.error(f"Plaid API Error: {e.body}")
            return None
        except Exception as e:
            logger.error(f"An unexpected error occurred fetching Plaid transactions: {e}")
            return None

    def get_item(self, access_token):
        """Returns Plaid Item metadata for an access token."""
        try:
            response = self.client.item_get(ItemGetRequest(access_token=access_token))
            return response.to_dict().get('item', {})
        except PlaidApiException as e:
            logger.error(f"Plaid item/get error: {e.body}")
            return {}
        except Exception as e:
            logger.error(f"Unexpected Plaid item/get error: {e}")
            return {}

    def get_accounts(self, access_token):
        """Returns account metadata for an access token."""
        try:
            response = self.client.accounts_get(AccountsGetRequest(access_token=access_token))
            return response.to_dict().get('accounts', [])
        except PlaidApiException as e:
            logger.error(f"Plaid accounts/get error: {e.body}")
            return []
        except Exception as e:
            logger.error(f"Unexpected Plaid accounts/get error: {e}")
            return []

    def get_institution_name(self, institution_id):
        """Returns a readable institution name for a Plaid institution id."""
        if not institution_id:
            return 'Unknown Institution'

        try:
            request = InstitutionsGetByIdRequest(
                institution_id=institution_id,
                country_codes=[CountryCode('US')]
            )
            response = self.client.institutions_get_by_id(request)
            return response.to_dict().get('institution', {}).get('name') or institution_id
        except PlaidApiException as e:
            logger.error(f"Plaid institution lookup error for {institution_id}: {e.body}")
            return institution_id
        except Exception as e:
            logger.error(f"Unexpected institution lookup error for {institution_id}: {e}")
            return institution_id

    def get_account_label_map(self, access_token):
        """Returns {account_id: friendly_label} for an access token."""
        item = self.get_item(access_token)
        institution_name = self.get_institution_name(item.get('institution_id'))
        accounts = self.get_accounts(access_token)

        labels = {}
        for account in accounts:
            account_id = account.get('account_id')
            if account_id:
                labels[account_id] = build_account_label(institution_name, account)
        return labels

    def create_transactions_link_token(self, client_user_id='local-user', redirect_uri=None):
        """Creates a Plaid Link token for the Transactions product only."""
        request_kwargs = {
            'products': [Products('transactions')],
            'client_name': 'Danny Bank Automation',
            'country_codes': [CountryCode('US')],
            'language': 'en',
            'user': LinkTokenCreateRequestUser(client_user_id=client_user_id),
        }
        if redirect_uri:
            request_kwargs['redirect_uri'] = redirect_uri

        response = self.client.link_token_create(LinkTokenCreateRequest(**request_kwargs))
        return response.to_dict().get('link_token')

    def exchange_public_token(self, public_token):
        """Exchanges a temporary public_token for a permanent access_token."""
        response = self.client.item_public_token_exchange(
            ItemPublicTokenExchangeRequest(public_token=public_token)
        )
        return response.to_dict()

    def describe_access_token(self, access_token):
        """Returns institution and account metadata for a Plaid access token."""
        item = self.get_item(access_token)
        institution_id = item.get('institution_id')
        institution_name = self.get_institution_name(institution_id)
        accounts = self.get_accounts(access_token)
        return {
            'institution_id': institution_id,
            'institution_name': institution_name,
            'accounts': accounts,
            'labels': [build_account_label(institution_name, account) for account in accounts],
        }
