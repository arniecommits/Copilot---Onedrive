"""
Standalone Script: Complete Bot to OneDrive/SharePoint Knowledge Source Mapper
Maps Copilot Studio bots to their OneDrive/SharePoint knowledge sources
Shows the complete relationship chain without importing other scripts

Requirements:
    pip install msal requests python-dotenv

Environment Variables (.env file):
    AZURE_CLIENT_ID=your-client-id
    AZURE_CLIENT_SECRET=your-client-secret
    AZURE_TENANT_ID=your-tenant-id
    DATAVERSE_URL=https://yourorg.crm.dynamics.com

Usage:
    python standalone_agent_knowledge_mapper.py
    python standalone_agent_knowledge_mapper.py --verbose
"""

import json
import os
import sys
import argparse
from typing import Dict, List, Optional
import msal
import requests
from dotenv import load_dotenv


class StandaloneKnowledgeMapper:
    """Complete standalone mapper for bot-to-knowledge-source relationships."""

    def __init__(self, client_id: str, client_secret: str, tenant_id: str, dataverse_url: str):
        """Initialize the mapper with credentials."""
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.dataverse_url = dataverse_url.rstrip('/')

        # Ensure https:// prefix
        if not self.dataverse_url.startswith('http'):
            self.dataverse_url = f"https://{self.dataverse_url}"

        self.authority = f"https://login.microsoftonline.com/{tenant_id}"
        self.scope = [f"{self.dataverse_url}/.default"]
        self.access_token = None

    def authenticate(self) -> str:
        """Acquire access token using MSAL."""
        print("Authenticating with Dataverse...")

        app = msal.ConfidentialClientApplication(
            self.client_id,
            authority=self.authority,
            client_credential=self.client_secret
        )

        result = app.acquire_token_for_client(scopes=self.scope)

        if "access_token" in result:
            self.access_token = result["access_token"]
            print("✓ Authentication successful\n")
            return self.access_token
        else:
            error_msg = result.get("error_description", result.get("error", "Unknown error"))
            raise Exception(f"Failed to acquire token: {error_msg}")

    def _make_request(self, endpoint: str, description: str = "") -> Dict:
        """Make a GET request to Dataverse API."""
        if not self.access_token:
            raise Exception("Not authenticated. Call authenticate() first.")

        url = f"{self.dataverse_url}/api/data/v9.2/{endpoint}"

        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Accept": "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0"
        }

        if description:
            print(f"{description}...")

        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.HTTPError as e:
            print(f"HTTP Error: {e}")
            print(f"Response: {response.text}")
            raise
        except requests.exceptions.RequestException as e:
            print(f"Request Error: {e}")
            raise

    def get_all_bots(self) -> List[Dict]:
        """Retrieve all Copilot Studio bots from Dataverse."""
        print("=" * 120)
        print("STEP 1: RETRIEVING ALL BOTS")
        print("=" * 120)
        print(f"API: GET {self.dataverse_url}/api/data/v9.2/bots")

        result = self._make_request("bots", "Fetching bots")
        bots = result.get("value", [])

        print(f"✓ Found {len(bots)} bot(s)")
        for i, bot in enumerate(bots, 1):
            print(f"  {i}. {bot.get('name', 'Unnamed')} (ID: {bot.get('botid')})")

        return bots

    def get_all_knowledge_sources(self) -> List[Dict]:
        """Retrieve all knowledge sources from dvtablesearchs table."""
        print("\n" + "=" * 120)
        print("STEP 2: RETRIEVING ALL KNOWLEDGE SOURCES")
        print("=" * 120)
        print(f"API: GET {self.dataverse_url}/api/data/v9.2/dvtablesearchs")

        endpoint = "dvtablesearchs?$select=dvtablesearchid,knowledgeconfig,name,appmoduleid"
        result = self._make_request(endpoint, "Fetching knowledge sources")
        knowledge_sources = result.get("value", [])

        print(f"✓ Found {len(knowledge_sources)} knowledge source(s)")

        # Count OneDrive/SharePoint sources
        onedrive_count = 0
        for ks in knowledge_sources:
            kconfig = ks.get('knowledgeconfig')
            if kconfig:
                try:
                    kconfig_obj = json.loads(kconfig) if isinstance(kconfig, str) else kconfig
                    if kconfig_obj.get('$kind') == 'IngestionBasedGraphSearchConfiguration':
                        if kconfig_obj.get('driveItems'):
                            onedrive_count += 1
                except:
                    pass

        print(f"  - OneDrive/SharePoint sources: {onedrive_count}")
        print(f"  - Other sources: {len(knowledge_sources) - onedrive_count}")

        return knowledge_sources

    def get_bot_components(self, bot_id: str) -> List[Dict]:
        """Retrieve all components for a specific bot."""
        endpoint = f"botcomponents?$filter=_parentbotid_value eq {bot_id}"
        result = self._make_request(endpoint)
        return result.get("value", [])

    def build_knowledge_source_maps(self, knowledge_sources: List[Dict], verbose: bool = False) -> tuple:
        """
        Build lookup maps for knowledge sources by ID and NAME.

        Returns:
            (ks_map_by_id, ks_map_by_name) tuple
        """
        ks_map_by_id = {}
        ks_map_by_name = {}

        for ks in knowledge_sources:
            ks_id = ks.get('dvtablesearchid')
            ks_name = ks.get('name', '')

            ks_map_by_id[ks_id] = ks
            if ks_name:
                ks_map_by_name[ks_name] = ks

            if verbose:
                kconfig = ks.get('knowledgeconfig')
                if kconfig:
                    try:
                        kconfig_obj = json.loads(kconfig) if isinstance(kconfig, str) else kconfig
                        if kconfig_obj.get('$kind') == 'IngestionBasedGraphSearchConfiguration':
                            drive_items = kconfig_obj.get('driveItems', [])
                            if drive_items:
                                print(f"\n  Knowledge Source: {ks_name}")
                                print(f"    GUID: {ks_id}")
                                print(f"    URLs:")
                                for item in drive_items:
                                    print(f"      - {item.get('displayName')}: {item.get('webUrl')}")
                    except:
                        pass

        return ks_map_by_id, ks_map_by_name

    def find_knowledge_source_references(self, component: Dict, ks_map_by_id: Dict, ks_map_by_name: Dict) -> List[tuple]:
        """
        Find knowledge source references in a component.

        Returns:
            List of (match_type, match_value, ks_object) tuples
        """
        matches = []

        # Get component data
        comp_data = component.get('data', '')
        comp_content = component.get('content', '')

        # Convert to searchable strings
        data_str = json.dumps(comp_data) if comp_data else ''
        content_str = json.dumps(comp_content) if comp_content else ''
        combined = data_str + content_str

        # METHOD 1: Search for GUIDs (older bots)
        for ks_id, ks_obj in ks_map_by_id.items():
            if ks_id and ks_id in combined:
                matches.append(('guid', ks_id, ks_obj))

        # METHOD 2: Search for NAMEs (newer bots)
        for ks_name, ks_obj in ks_map_by_name.items():
            if ks_name and ks_name in combined:
                matches.append(('name', ks_name, ks_obj))

        return matches

    def extract_urls_from_knowledge_source(self, ks_obj: Dict) -> List[Dict]:
        """
        Extract OneDrive/SharePoint URLs from a knowledge source.

        Returns:
            List of {type, name, url, knowledge_source_id, knowledge_source_name} dicts
        """
        results = []
        kconfig = ks_obj.get('knowledgeconfig')

        if not kconfig:
            return results

        try:
            kconfig_obj = json.loads(kconfig) if isinstance(kconfig, str) else kconfig
            kind = kconfig_obj.get('$kind')

            # OneDrive/SharePoint sources
            if kind == 'IngestionBasedGraphSearchConfiguration':
                drive_items = kconfig_obj.get('driveItems', [])
                for item in drive_items:
                    web_url = item.get('webUrl', '')
                    display_name = item.get('displayName', 'N/A')

                    # Categorize by URL pattern
                    if '-my.sharepoint.com' in web_url.lower() or '/personal/' in web_url.lower():
                        source_type = 'OneDrive'
                    elif 'sharepoint.com' in web_url.lower():
                        source_type = 'SharePoint'
                    else:
                        source_type = 'Unknown'

                    results.append({
                        'type': source_type,
                        'name': display_name,
                        'url': web_url,
                        'knowledge_source_id': ks_obj.get('dvtablesearchid'),
                        'knowledge_source_name': ks_obj.get('name', 'N/A')
                    })

            # Dataverse sources
            elif kind == 'SqlFederatedTableSearchConfiguration':
                results.append({
                    'type': 'Dataverse',
                    'server': kconfig_obj.get('sqlServerName', 'N/A'),
                    'database': kconfig_obj.get('sqlDbName', 'N/A'),
                    'knowledge_source_id': ks_obj.get('dvtablesearchid'),
                    'knowledge_source_name': ks_obj.get('name', 'N/A')
                })

        except Exception as e:
            print(f"    [Warning] Could not parse knowledge config: {e}")

        return results

    def map_all_bots_to_knowledge_sources(self, verbose: bool = False) -> Dict:
        """
        Complete mapping process for all bots.

        Returns:
            Results dictionary with mappings and categorized agents
        """
        # Step 1: Get all bots
        bots = self.get_all_bots()

        if not bots:
            print("No bots found!")
            return {}

        # Step 2: Get all knowledge sources
        knowledge_sources = self.get_all_knowledge_sources()

        # Build lookup maps
        ks_map_by_id, ks_map_by_name = self.build_knowledge_source_maps(knowledge_sources, verbose)

        print(f"  - Unique knowledge source names: {len(ks_map_by_name)}")

        # Results tracking
        results = {
            'detailed_mappings': [],
            'agents_with_onedrive': [],
            'agents_with_sharepoint': [],
            'agents_with_other_sources': [],
            'agents_without_sources': []
        }

        # Step 3: Process each bot
        print("\n" + "=" * 120)
        print("STEP 3: MAPPING BOTS TO KNOWLEDGE SOURCES")
        print("=" * 120)

        for bot_num, bot in enumerate(bots, 1):
            bot_id = bot.get('botid')
            bot_name = bot.get('name', 'Unnamed Bot')

            print(f"\n{'-' * 120}")
            print(f"BOT {bot_num}/{len(bots)}: {bot_name}")
            print(f"Bot ID: {bot_id}")
            print(f"{'-' * 120}")

            try:
                # Get bot components
                print(f"API: GET {self.dataverse_url}/api/data/v9.2/botcomponents?$filter=_parentbotid_value eq {bot_id}")
                components = self.get_bot_components(bot_id)
                print(f"✓ Retrieved {len(components)} component(s)")

                if not components:
                    results['agents_without_sources'].append({
                        'name': bot_name,
                        'id': bot_id,
                        'reason': 'No components found'
                    })
                    print("  └─ No components found")
                    continue

                # Search for knowledge source references
                print(f"\nSearching for knowledge source references (GUIDs and NAMEs)...")

                found_ks_by_id = {}  # ks_id -> list of component references

                for comp_idx, component in enumerate(components, 1):
                    comp_id = component.get('botcomponentid')
                    comp_type = component.get('componenttype')
                    comp_name = component.get('name', 'N/A')

                    matches = self.find_knowledge_source_references(component, ks_map_by_id, ks_map_by_name)

                    if matches and verbose:
                        print(f"\n  Component #{comp_idx}: {comp_name} (Type: {comp_type})")
                        for match_type, match_value, _ in matches:
                            print(f"    ✓ Found {match_type.upper()}: {match_value}")

                    # Record matches
                    for match_type, match_value, ks_obj in matches:
                        ks_id = ks_obj.get('dvtablesearchid')
                        if ks_id not in found_ks_by_id:
                            found_ks_by_id[ks_id] = []
                        found_ks_by_id[ks_id].append({
                            'component_id': comp_id,
                            'component_type': comp_type,
                            'component_name': comp_name,
                            'match_type': match_type,
                            'match_value': match_value
                        })

                if not found_ks_by_id:
                    results['agents_without_sources'].append({
                        'name': bot_name,
                        'id': bot_id,
                        'reason': 'No knowledge source references found'
                    })
                    print("  └─ No knowledge source references found")
                    continue

                print(f"\n✓ Found {len(found_ks_by_id)} unique knowledge source reference(s)")
                print(f"\nExtracting URLs...")

                # Extract URLs from found knowledge sources
                onedrive_urls = []
                sharepoint_urls = []
                other_sources = []

                for ks_id, component_refs in found_ks_by_id.items():
                    ks_obj = ks_map_by_id.get(ks_id)
                    if not ks_obj:
                        continue

                    ks_name = ks_obj.get('name', 'N/A')

                    print(f"\n  Knowledge Source: {ks_name}")
                    print(f"    GUID: {ks_id}")
                    print(f"    Referenced by {len(component_refs)} component(s):")
                    for ref in component_refs:
                        match_info = f"via {ref['match_type'].upper()}: {ref['match_value']}"
                        print(f"      - {ref['component_name']} (Type: {ref['component_type']}) {match_info}")

                    # Extract URLs
                    urls = self.extract_urls_from_knowledge_source(ks_obj)

                    for url_info in urls:
                        if url_info['type'] == 'OneDrive':
                            print(f"      └─ [OneDrive] {url_info['name']}")
                            print(f"         URL: {url_info['url']}")
                            onedrive_urls.append(url_info)
                        elif url_info['type'] == 'SharePoint':
                            print(f"      └─ [SharePoint] {url_info['name']}")
                            print(f"         URL: {url_info['url']}")
                            sharepoint_urls.append(url_info)
                        elif url_info['type'] == 'Dataverse':
                            print(f"      └─ [Dataverse] {url_info['server']}/{url_info['database']}")
                            other_sources.append(url_info)

                        # Store complete mapping
                        results['detailed_mappings'].append({
                            'bot_name': bot_name,
                            'bot_id': bot_id,
                            'component_references': component_refs,
                            'knowledge_source_id': ks_id,
                            'knowledge_source_name': ks_name,
                            **url_info
                        })

                # Categorize bot
                agent_info = {
                    'name': bot_name,
                    'id': bot_id,
                    'created': bot.get('createdon', 'N/A'),
                    'modified': bot.get('modifiedon', 'N/A')
                }

                if onedrive_urls:
                    unique_onedrive = {item['url']: item for item in onedrive_urls}.values()
                    agent_info['onedrive_sources'] = list(unique_onedrive)
                    results['agents_with_onedrive'].append(agent_info.copy())
                    print(f"\n✓ RESULT: {len(unique_onedrive)} OneDrive source(s)")

                if sharepoint_urls:
                    unique_sharepoint = {item['url']: item for item in sharepoint_urls}.values()
                    agent_info['sharepoint_sources'] = list(unique_sharepoint)
                    results['agents_with_sharepoint'].append(agent_info.copy())
                    print(f"✓ RESULT: {len(unique_sharepoint)} SharePoint source(s)")

                if other_sources:
                    agent_info['other_sources'] = other_sources
                    results['agents_with_other_sources'].append(agent_info.copy())
                    print(f"✓ RESULT: {len(other_sources)} other source(s)")

                if not onedrive_urls and not sharepoint_urls and not other_sources:
                    results['agents_without_sources'].append(agent_info)
                    print(f"✗ RESULT: No OneDrive/SharePoint sources")

            except Exception as e:
                print(f"\n✗ ERROR: {e}")
                import traceback
                traceback.print_exc()
                results['agents_without_sources'].append({
                    'name': bot_name,
                    'id': bot_id,
                    'error': str(e)
                })

        return results

    def print_summary(self, results: Dict):
        """Print summary report."""
        print("\n" + "=" * 120)
        print("STEP 4: SUMMARY REPORT")
        print("=" * 120)

        total = (len(results['agents_with_onedrive']) +
                len(results['agents_with_sharepoint']) +
                len(results['agents_with_other_sources']) +
                len(results['agents_without_sources']))

        print(f"\nTotal Bots Analyzed: {total}")
        print(f"  ✓ Bots with OneDrive access: {len(results['agents_with_onedrive'])}")
        print(f"  ✓ Bots with SharePoint access: {len(results['agents_with_sharepoint'])}")
        print(f"  ✓ Bots with other sources: {len(results['agents_with_other_sources'])}")
        print(f"  ✗ Bots without sources: {len(results['agents_without_sources'])}")

        # Relationship chains
        if results['detailed_mappings']:
            print("\n" + "=" * 120)
            print("COMPLETE RELATIONSHIP CHAINS")
            print("=" * 120)

            current_bot = None
            for mapping in results['detailed_mappings']:
                if mapping['bot_name'] != current_bot:
                    current_bot = mapping['bot_name']
                    print(f"\n┌─ BOT: {current_bot}")
                    print(f"│  Bot ID: {mapping['bot_id']}")

                print(f"│")
                for comp_ref in mapping['component_references']:
                    match_info = f"{comp_ref['match_type'].upper()}: {comp_ref['match_value']}"
                    print(f"├─── COMPONENT: {comp_ref['component_name']} (Type: {comp_ref['component_type']})")
                    print(f"│    Match: {match_info}")

                print(f"│")
                print(f"├───── KNOWLEDGE SOURCE: {mapping['knowledge_source_name']}")
                print(f"│      GUID: {mapping['knowledge_source_id']}")
                print(f"│")

                if mapping.get('url'):
                    print(f"└─────── [{mapping['type']}] {mapping['name']}")
                    print(f"         URL: {mapping['url']}")
                elif mapping.get('server'):
                    print(f"└─────── [Dataverse] {mapping['server']}/{mapping['database']}")
                print()

        # Detailed reports
        if results['agents_with_onedrive']:
            print("\n" + "=" * 120)
            print("BOTS WITH ONEDRIVE ACCESS")
            print("=" * 120)
            for agent in results['agents_with_onedrive']:
                print(f"\n{agent['name']}")
                print(f"  Bot ID: {agent['id']}")
                print(f"  OneDrive Sources:")
                for source in agent.get('onedrive_sources', []):
                    print(f"    • {source['name']}")
                    print(f"      URL: {source['url']}")
                    print(f"      Knowledge Source: {source.get('knowledge_source_name', 'N/A')}")

        if results['agents_with_sharepoint']:
            print("\n" + "=" * 120)
            print("BOTS WITH SHAREPOINT ACCESS")
            print("=" * 120)
            for agent in results['agents_with_sharepoint']:
                print(f"\n{agent['name']}")
                print(f"  Bot ID: {agent['id']}")
                print(f"  SharePoint Sources:")
                for source in agent.get('sharepoint_sources', []):
                    print(f"    • {source['name']}")
                    print(f"      URL: {source['url']}")
                    print(f"      Knowledge Source: {source.get('knowledge_source_name', 'N/A')}")

    def export_results(self, results: Dict, filename: str = "standalone_mapping_results.json"):
        """Export results to JSON file."""
        print("\n" + "=" * 120)
        print("EXPORTING RESULTS")
        print("=" * 120)
        print(f"Writing to '{filename}'...")

        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)

        print(f"✓ Export complete")


def main():
    """Main execution function."""
    parser = argparse.ArgumentParser(
        description='Standalone Bot-to-Knowledge-Source Mapper',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
This script maps Copilot Studio bots to their OneDrive/SharePoint knowledge sources.
It shows the complete relationship chain:
  Bot → Component → Reference (GUID/NAME) → Knowledge Source → URL

Examples:
  python standalone_agent_knowledge_mapper.py
  python standalone_agent_knowledge_mapper.py --verbose
  python standalone_agent_knowledge_mapper.py -v -o my_results.json

The script handles TWO types of references:
  1. GUID-based (older bots): Component contains dvtablesearchid GUID
  2. NAME-based (newer bots): Component contains knowledge source name (e.g., skillConfiguration)
        """
    )
    parser.add_argument('-v', '--verbose', action='store_true',
                       help='Show detailed matching information for each component')
    parser.add_argument('-o', '--output', default='standalone_mapping_results.json',
                       help='Output JSON filename (default: standalone_mapping_results.json)')

    args = parser.parse_args()

    # Load environment variables
    load_dotenv()

    CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
    CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
    TENANT_ID = os.getenv("AZURE_TENANT_ID")
    DATAVERSE_URL = os.getenv("DATAVERSE_URL")

    if not all([CLIENT_ID, CLIENT_SECRET, TENANT_ID, DATAVERSE_URL]):
        print("=" * 120)
        print("ERROR: Missing required environment variables")
        print("=" * 120)
        print("\nRequired in .env file:")
        print("  AZURE_CLIENT_ID")
        print("  AZURE_CLIENT_SECRET")
        print("  AZURE_TENANT_ID")
        print("  DATAVERSE_URL")
        print("\nExample .env file:")
        print("  AZURE_CLIENT_ID=12345678-1234-1234-1234-123456789abc")
        print("  AZURE_CLIENT_SECRET=your-secret-here")
        print("  AZURE_TENANT_ID=87654321-4321-4321-4321-cba987654321")
        print("  DATAVERSE_URL=https://yourorg.crm.dynamics.com")
        sys.exit(1)

    try:
        print("=" * 120)
        print("STANDALONE BOT-TO-KNOWLEDGE-SOURCE MAPPER")
        print("=" * 120)
        print(f"\nMode: {'VERBOSE' if args.verbose else 'NORMAL'}")
        print(f"Dataverse URL: {DATAVERSE_URL}")
        print(f"Output File: {args.output}\n")

        # Initialize mapper
        mapper = StandaloneKnowledgeMapper(CLIENT_ID, CLIENT_SECRET, TENANT_ID, DATAVERSE_URL)

        # Authenticate
        mapper.authenticate()

        # Map all bots to knowledge sources
        results = mapper.map_all_bots_to_knowledge_sources(verbose=args.verbose)

        # Print summary
        mapper.print_summary(results)

        # Export results
        mapper.export_results(results, args.output)

        print("\n" + "=" * 120)
        print("COMPLETE!")
        print("=" * 120)
        print(f"\nOutput file: {args.output}")
        print("This file contains the complete relationship mappings in JSON format.")

    except Exception as e:
        print("\n" + "=" * 120)
        print(f"ERROR: {e}")
        print("=" * 120)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
