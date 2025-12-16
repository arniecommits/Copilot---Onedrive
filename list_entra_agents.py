"""
Microsoft Entra ID Agent Identity Listing Script
Uses Microsoft Graph API (Beta) to list agent identities
API Reference: https://learn.microsoft.com/en-us/graph/api/agentidentity-list?view=graph-rest-beta
"""

import requests
import json
from typing import Optional, Dict, List
import msal
import os
from dotenv import load_dotenv


class EntraAgentLister:
    """Client for listing Entra ID Agent Identities using Microsoft Graph API."""

    def __init__(self, client_id: str, client_secret: str, tenant_id: str):
        """
        Initialize the Entra Agent Lister.

        Args:
            client_id: Azure AD application (client) ID
            client_secret: Client secret for the application
            tenant_id: Azure AD tenant ID
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.authority = f"https://login.microsoftonline.com/{tenant_id}"
        self.scope = ["https://graph.microsoft.com/.default"]
        self.graph_endpoint = "https://graph.microsoft.com/beta/$batch"
        self.access_token = None

    def get_access_token(self) -> str:
        """
        Acquire access token using MSAL (Microsoft Authentication Library).

        Returns:
            Access token string
        """
        app = msal.ConfidentialClientApplication(
            self.client_id,
            authority=self.authority,
            client_credential=self.client_secret
        )

        result = app.acquire_token_for_client(scopes=self.scope)

        if "access_token" in result:
            self.access_token = result["access_token"]
            return self.access_token
        else:
            error_msg = result.get("error_description", result.get("error", "Unknown error"))
            raise Exception(f"Failed to acquire token: {error_msg}")

    def list_agents(self,
                   select: Optional[List[str]] = None,
                   filter_query: Optional[str] = None,
                   orderby: Optional[str] = None,
                   top: Optional[int] = None,
                   search: Optional[str] = None) -> Dict:
        """
        List all agent identities with optional query parameters using batch API.

        Args:
            select: List of properties to select (e.g., ['id', 'displayName'])
            filter_query: Additional OData filter string (will be combined with base filter)
            orderby: Property to order by (e.g., "displayName")
            top: Maximum number of results to return
            search: Search query string

        Returns:
            Dictionary containing the API response with agent identities
        """
        if not self.access_token:
            self.get_access_token()

        # Base filter for agent identities
        base_filter = "(isof('microsoft.graph.agentIdentity') OR (tags/any(p:startswith(p, 'power-virtual-agents-')) OR tags/any(p:p eq 'AgenticInstance')))"

        # Combine with additional filter if provided
        if filter_query:
            combined_filter = f"({base_filter}) and ({filter_query})"
        else:
            combined_filter = base_filter

        # Build URL with query parameters
        url = "/servicePrincipals/?$count=true&$filter=" + combined_filter

        if select:
            url += "&$select=" + ",".join(select)
        if top:
            url += f"&$top={top}"
        if search:
            url += f"&$search={search}"

        # Create batch request
        batch_request = {
            "requests": [
                {
                    "id": "1",
                    "method": "GET",
                    "url": url,
                    "headers": {
                        "ConsistencyLevel": "eventual"
                    }
                }
            ]
        }

        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }

        try:
            response = requests.post(
                self.graph_endpoint,
                headers=headers,
                json=batch_request
            )
            response.raise_for_status()
            batch_response = response.json()

            # Extract the actual response from batch response
            if "responses" in batch_response and len(batch_response["responses"]) > 0:
                inner_response = batch_response["responses"][0]
                if inner_response.get("status") == 200:
                    return inner_response.get("body", {})
                else:
                    error_body = inner_response.get("body", {})
                    raise Exception(f"Batch request failed with status {inner_response.get('status')}: {error_body}")
            else:
                raise Exception("Invalid batch response structure")

        except requests.exceptions.HTTPError as e:
            print(f"HTTP Error: {e}")
            print(f"Response: {response.text}")
            raise
        except requests.exceptions.RequestException as e:
            print(f"Request Error: {e}")
            raise

    def print_agents(self, agents_response: Dict):
        """
        Pretty print agent identities from the API response.

        Args:
            agents_response: Dictionary containing the API response
        """
        agents = agents_response.get("value", [])

        if not agents:
            print("No agent identities found.")
            return

        print(f"\nFound {len(agents)} agent identities:\n")
        print("=" * 80)

        for idx, agent in enumerate(agents, 1):
            print(f"\nAgent #{idx}")
            print("-" * 80)
            print(f"  ID:                    {agent.get('id', 'N/A')}")
            print(f"  Display Name:          {agent.get('displayName', 'N/A')}")
            print(f"  Created DateTime:      {agent.get('createdDateTime', 'N/A')}")
            print(f"  Created By App ID:     {agent.get('createdByAppId', 'N/A')}")
            print(f"  Blueprint ID:          {agent.get('agentIdentityBlueprintId', 'N/A')}")
            print(f"  Account Enabled:       {agent.get('accountEnabled', 'N/A')}")
            print(f"  Service Principal Type: {agent.get('servicePrincipalType', 'N/A')}")
            print(f"  Disabled by Microsoft: {agent.get('disabledByMicrosoftStatus', 'N/A')}")

            tags = agent.get('tags', [])
            if tags:
                print(f"  Tags:                  {', '.join(tags)}")
            else:
                print(f"  Tags:                  None")

        print("\n" + "=" * 80)


def main():
    """
    Main function to demonstrate usage of the EntraAgentLister.

    Required Azure AD App Registration:
    1. Register an application in Azure Portal
    2. Add API permission: Microsoft Graph > Application.Read.All (Application permission)
    3. Grant admin consent for the permission
    4. Create a client secret
    5. Set credentials in .env file
    """

    # Load environment variables from .env file
    load_dotenv()

    # Configuration - Read from environment variables
    CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
    CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
    TENANT_ID = os.getenv("AZURE_TENANT_ID")

    # Validate configuration
    if not CLIENT_ID or not CLIENT_SECRET or not TENANT_ID:
        print("ERROR: Please configure Azure AD credentials in the .env file")
        print("\nRequired environment variables in .env:")
        print("  AZURE_CLIENT_ID=your-application-client-id")
        print("  AZURE_CLIENT_SECRET=your-client-secret-value")
        print("  AZURE_TENANT_ID=your-directory-tenant-id")
        print("\nRequired steps:")
        print("1. Register an application in Azure Portal (https://portal.azure.com)")
        print("2. Navigate to 'App registrations' and create a new registration")
        print("3. Copy the Application (client) ID and Directory (tenant) ID")
        print("4. Go to 'Certificates & secrets' and create a new client secret")
        print("5. Go to 'API permissions' and add:")
        print("   - Microsoft Graph > Application permissions > Application.Read.All")
        print("6. Click 'Grant admin consent'")
        print("7. Add credentials to .env file")
        print("\nRequired Permission: Application.Read.All (Application permission)")
        return

    try:
        # Initialize the client
        print("Initializing Entra Agent Lister...")
        lister = EntraAgentLister(CLIENT_ID, CLIENT_SECRET, TENANT_ID)

        # Acquire access token
        print("Acquiring access token...")
        lister.get_access_token()
        print("[OK] Successfully authenticated")

        # Example 1: List all agents
        print("\n[Example 1] Listing all agent identities...")
        agents = lister.list_agents()
        lister.print_agents(agents)

        # Example 2: List only enabled agents
        print("\n[Example 2] Listing only enabled agents...")
        enabled_agents = lister.list_agents(filter_query="accountEnabled eq true")
        lister.print_agents(enabled_agents)

        # Example 3: List agents with selected properties
        print("\n[Example 3] Listing agents with selected properties...")
        selected_agents = lister.list_agents(
            select=["id", "displayName", "accountEnabled", "createdDateTime"]
        )
        lister.print_agents(selected_agents)

        # Export to JSON file
        print("\nExporting all agents to 'entra_agents.json'...")
        with open('entra_agents.json', 'w', encoding='utf-8') as f:
            json.dump(agents, f, indent=2, ensure_ascii=False)
        print("[OK] Export complete")

    except Exception as e:
        print(f"\n[ERROR] {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
