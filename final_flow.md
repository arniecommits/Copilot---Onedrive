# Final Flow: Bot to OneDrive URL Mapping

This guide shows the complete flow to map Copilot Studio bots to their OneDrive/SharePoint knowledge sources.

## Overview: The Relationship Chain

```
BOT → BOT COMPONENTS → KNOWLEDGE SOURCE REFERENCE → KNOWLEDGE SOURCE → OneDrive/SharePoint URLs
```

**Key Insight:** The relationship is INDIRECT and uses either:
- **NAME matching** (newer bots) - Component contains knowledge source name
- **GUID matching** (older bots) - Component contains knowledge source ID

---

## Step 1: Get Access Token

**API Call:**
```http
POST https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token

Body:
  client_id={CLIENT_ID}
  client_secret={CLIENT_SECRET}
  scope=https://{ORG}.crm.dynamics.com/.default
  grant_type=client_credentials
```

**Response:**
```json
{
  "access_token": "eyJ0eXAiOiJKV1Qi...",
  "expires_in": 3599
}
```

**What to extract:** Save `access_token` for all subsequent API calls.

---

## Step 2: Get All Bots

**API Call:**
```http
GET https://{ORG}.crm.dynamics.com/api/data/v9.2/bots
Authorization: Bearer {access_token}
```

**Response:**
```json
{
  "value": [
    {
      "botid": "b4dc0143-cbda-f011-8543-000d3a5ae560",
      "name": "Ondrive Agent",
      "createdon": "2025-12-05T08:59:40Z"
    },
    {
      "botid": "57c1a9b2-b8d1-f011-8543-000d3a5ae560",
      "name": "Customer Service Bot",
      "createdon": "2025-11-20T10:30:15Z"
    }
  ]
}
```

**What to extract:** For each bot, save:
- `botid` - This is what we'll use in Step 4
- `name` - Bot display name

---

## Step 3: Get All Knowledge Sources

**API Call:**
```http
GET https://{ORG}.crm.dynamics.com/api/data/v9.2/dvtablesearchs?$select=dvtablesearchid,name,knowledgeconfig
Authorization: Bearer {access_token}
```

**Response:**
```json
{
  "value": [
    {
      "dvtablesearchid": "a0678757-ce83-49c1-b7db-25c3d6474547",
      "name": "Heydata_6DG3Bf6BjJbzQ3uTmcNr_",
      "knowledgeconfig": "{\"$kind\":\"IngestionBasedGraphSearchConfiguration\",\"driveItems\":[{\"displayName\":\"README Files\",\"webUrl\":\"https://contoso-my.sharepoint.com/personal/user/Documents/README\"}]}"
    },
    {
      "dvtablesearchid": "bb678757-ce83-49c1-b7db-25c3d6474547",
      "name": "Customer Files",
      "knowledgeconfig": "{\"$kind\":\"IngestionBasedGraphSearchConfiguration\",\"driveItems\":[{\"displayName\":\"Sales Folder\",\"webUrl\":\"https://contoso.sharepoint.com/sites/sales/documents\"}]}"
    }
  ]
}
```

**What to extract:** Build TWO lookup dictionaries:

```python
# Dictionary 1: Lookup by GUID (for older bots)
ks_by_id = {
    "a0678757-ce83-49c1-b7db-25c3d6474547": {
        "name": "Heydata_6DG3Bf6BjJbzQ3uTmcNr_",
        "knowledgeconfig": "{...JSON with URLs...}"
    },
    "bb678757-ce83-49c1-b7db-25c3d6474547": {
        "name": "Customer Files",
        "knowledgeconfig": "{...JSON with URLs...}"
    }
}

# Dictionary 2: Lookup by NAME (for newer bots)
ks_by_name = {
    "Heydata_6DG3Bf6BjJbzQ3uTmcNr_": {
        "dvtablesearchid": "a0678757-ce83-49c1-b7db-25c3d6474547",
        "knowledgeconfig": "{...JSON with URLs...}"
    },
    "Customer Files": {
        "dvtablesearchid": "bb678757-ce83-49c1-b7db-25c3d6474547",
        "knowledgeconfig": "{...JSON with URLs...}"
    }
}
```

---

## Step 4: Get Bot Components (For Each Bot)

**API Call:** (Using botid from Step 2)
```http
GET https://{ORG}.crm.dynamics.com/api/data/v9.2/botcomponents?$filter=_parentbotid_value eq b4dc0143-cbda-f011-8543-000d3a5ae560
Authorization: Bearer {access_token}
```

**Response:**
```json
{
  "value": [
    {
      "botcomponentid": "70b1062a-c2bb-4adf-aa2b-2b50a9498f27",
      "componenttype": 15,
      "name": "Ondrive Agent",
      "data": "kind: GptComponentMetadata\ndisplayName: Ondrive Agent"
    },
    {
      "botcomponentid": "215b8870-36f5-426e-b341-8abd0a287b5f",
      "componenttype": 16,
      "name": "Hey data",
      "data": "kind: KnowledgeSourceConfiguration\nsource:\n  kind: FederatedStructuredSearchSource\n  skillConfiguration: Heydata_6DG3Bf6BjJbzQ3uTmcNr_"
    }
  ]
}
```

**What to look for in `data` field:**
- **NAME reference**: `skillConfiguration: Heydata_6DG3Bf6BjJbzQ3uTmcNr_`
- **GUID reference**: `a0678757-ce83-49c1-b7db-25c3d6474547`

---

## Step 5: Match Components to Knowledge Sources

### Example 1: NAME-Based Matching (Newer Bots)

**From Step 4 Component:**
```yaml
data: "kind: KnowledgeSourceConfiguration
source:
  kind: FederatedStructuredSearchSource
  skillConfiguration: Heydata_6DG3Bf6BjJbzQ3uTmcNr_"
```

**Matching Logic:**
1. Extract `skillConfiguration` value: `Heydata_6DG3Bf6BjJbzQ3uTmcNr_`
2. Search in `ks_by_name` dictionary → FOUND!
3. Get knowledge source object with `dvtablesearchid`: `a0678757-ce83-49c1-b7db-25c3d6474547`

**Visual Flow:**
```
Component Data → "skillConfiguration: Heydata_6DG3Bf6BjJbzQ3uTmcNr_"
                                    ↓
                          Search in ks_by_name{}
                                    ↓
                    Found knowledge source object!
                                    ↓
                    dvtablesearchid: a0678757-ce83-49c1-b7db-25c3d6474547
                    knowledgeconfig: {...}
```

---

### Example 2: GUID-Based Matching (Older Bots)

**From Step 4 Component:**
```json
{
  "data": "{\"knowledgeSourceId\":\"bb678757-ce83-49c1-b7db-25c3d6474547\",\"settings\":{...}}"
}
```

**Matching Logic:**
1. Extract GUID pattern: `bb678757-ce83-49c1-b7db-25c3d6474547`
2. Search in `ks_by_id` dictionary → FOUND!
3. Get knowledge source object

**Visual Flow:**
```
Component Data → "knowledgeSourceId":"bb678757-ce83-49c1-b7db-25c3d6474547"
                                    ↓
                          Search in ks_by_id{}
                                    ↓
                    Found knowledge source object!
                                    ↓
                    name: Customer Files
                    knowledgeconfig: {...}
```

---

## Step 6: Extract OneDrive/SharePoint URLs

**From Knowledge Source `knowledgeconfig` field:**
```json
{
  "$kind": "IngestionBasedGraphSearchConfiguration",
  "driveItems": [
    {
      "displayName": "README Files",
      "webUrl": "https://contoso-my.sharepoint.com/personal/user/Documents/README"
    }
  ]
}
```

**Extraction Logic:**
```python
import json

kconfig = json.loads(knowledge_source['knowledgeconfig'])

if kconfig.get('$kind') == 'IngestionBasedGraphSearchConfiguration':
    for item in kconfig.get('driveItems', []):
        url = item.get('webUrl')
        display_name = item.get('displayName')

        # Categorize
        if '-my.sharepoint.com' in url or '/personal/' in url:
            source_type = 'OneDrive'
        elif 'sharepoint.com' in url:
            source_type = 'SharePoint'

        print(f"{source_type}: {display_name} - {url}")
```

---

## Complete Example: End-to-End

### Input:
- Bot: "Ondrive Agent" (botid: `b4dc0143-cbda-f011-8543-000d3a5ae560`)

### Execution Flow:

**1. Get Bot Components** (Step 4)
```json
{
  "name": "Hey data",
  "data": "skillConfiguration: Heydata_6DG3Bf6BjJbzQ3uTmcNr_"
}
```

**2. Match to Knowledge Source** (Step 5)
```
Search for "Heydata_6DG3Bf6BjJbzQ3uTmcNr_" in ks_by_name{}
    ↓
Found! dvtablesearchid: a0678757-ce83-49c1-b7db-25c3d6474547
```

**3. Parse knowledgeconfig** (Step 6)
```json
{
  "$kind": "IngestionBasedGraphSearchConfiguration",
  "driveItems": [
    {
      "displayName": "README Files",
      "webUrl": "https://contoso-my.sharepoint.com/personal/user/Documents/README"
    }
  ]
}
```

**4. Final Output:**
```
Bot: Ondrive Agent
  └─ Knowledge Source: Heydata_6DG3Bf6BjJbzQ3uTmcNr_ (via NAME reference)
      └─ OneDrive: README Files
          └─ URL: https://contoso-my.sharepoint.com/personal/user/Documents/README
```

---

## Quick Reference: Two Matching Methods

| Method | Bot Type | What to Look For | Dictionary to Use |
|--------|----------|------------------|-------------------|
| **NAME** | Newer (2024+) | `skillConfiguration: Heydata_...` | `ks_by_name{}` |
| **GUID** | Older (pre-2024) | `"knowledgeSourceId":"a0678757-..."` | `ks_by_id{}` |

---

## Python Implementation Summary

```python
# STEP 1: Authenticate
access_token = get_access_token()

# STEP 2: Get all bots
bots = requests.get(f"{dataverse_url}/api/data/v9.2/bots",
                    headers={"Authorization": f"Bearer {access_token}"}).json()

# STEP 3: Get all knowledge sources and build lookup maps
knowledge_sources = requests.get(f"{dataverse_url}/api/data/v9.2/dvtablesearchs",
                                 headers={"Authorization": f"Bearer {access_token}"}).json()

ks_by_id = {}     # {dvtablesearchid: ks_object}
ks_by_name = {}   # {name: ks_object}

for ks in knowledge_sources['value']:
    ks_by_id[ks['dvtablesearchid']] = ks
    ks_by_name[ks['name']] = ks

# STEP 4-6: For each bot, find knowledge sources
for bot in bots['value']:
    bot_id = bot['botid']
    bot_name = bot['name']

    # Get bot components
    components = requests.get(
        f"{dataverse_url}/api/data/v9.2/botcomponents?$filter=_parentbotid_value eq {bot_id}",
        headers={"Authorization": f"Bearer {access_token}"}
    ).json()

    # Search for knowledge source references
    for component in components['value']:
        data_str = str(component.get('data', ''))

        # METHOD 1: Search for NAME references
        for ks_name, ks_obj in ks_by_name.items():
            if ks_name in data_str:
                print(f"✓ Found NAME reference: {ks_name}")
                extract_urls(ks_obj)

        # METHOD 2: Search for GUID references
        for ks_id, ks_obj in ks_by_id.items():
            if ks_id in data_str:
                print(f"✓ Found GUID reference: {ks_id}")
                extract_urls(ks_obj)

def extract_urls(knowledge_source):
    kconfig = json.loads(knowledge_source['knowledgeconfig'])
    if kconfig.get('$kind') == 'IngestionBasedGraphSearchConfiguration':
        for item in kconfig.get('driveItems', []):
            print(f"  URL: {item.get('webUrl')}")
```

---

## Key Takeaways

1. **Build lookup maps FIRST** (Step 3) - This makes matching fast
2. **Use BOTH dictionaries** - ks_by_id AND ks_by_name
3. **Search component data** for both names and GUIDs
4. **Parse knowledgeconfig JSON** to extract actual URLs
5. **The relationship is INDIRECT** - components contain references, not direct links

---

## Troubleshooting

**Problem:** Bot shows no knowledge sources

**Solution:** Check BOTH matching methods:
1. Print all component `data` fields
2. Print all knowledge source `name` values
3. Check if any names appear in component data
4. Check if any GUIDs appear in component data

**Script:** Use `check_guid_matches.py` to diagnose
