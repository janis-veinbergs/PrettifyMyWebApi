(async function () {
    async function odataFetch(url) {
        const response = await fetch(url, { headers: { 'Prefer': 'odata.include-annotations="*"', 'Cache-Control': 'no-cache' } });

        if (!response.ok) {
            const errorText = await response.text();
            throw `${response.status} - ${errorText}`;
        }

        return await response.json();
    }

    function normalizeGuid(guid) {
        if (!guid) {
            return null;
        }

        const normalized = guid.replace('{', '').replace('}', '').trim();
        if (/^[0-9a-fA-F-]{36}$/.test(normalized) === false) {
            return null;
        }

        return normalized;
    }

    function getApiPrefixPath() {
        const parts = window.location.pathname.split('/').filter(Boolean);
        if (parts.length <= 1) {
            return '';
        }

        if (parts[0].toLowerCase() === 'main.aspx') {
            return '';
        }

        return '/' + parts[0];
    }

    function getVersionCandidates(preferredVersion) {
        const versions = [];

        if (preferredVersion) {
            versions.push(preferredVersion);
        }

        ['9.2', '9.1', '9.0'].forEach(v => {
            if (versions.includes(v) === false) {
                versions.push(v);
            }
        });

        return versions;
    }

    async function getWebApiBaseUrl(entityLogicalName, preferredVersion) {
        const prefixPath = getApiPrefixPath();
        const versions = getVersionCandidates(preferredVersion);

        for (let version of versions) {
            const apiUrl = `${window.location.origin}${prefixPath}/api/data/v${version}/`;
            const requestUrl = `${apiUrl}EntityDefinitions?$select=EntitySetName&$filter=(LogicalName eq %27${entityLogicalName}%27)`;

            try {
                const result = await odataFetch(requestUrl);
                if (result?.value?.length > 0) {
                    const pluralName = result.value[0].EntitySetName;
                    return `${apiUrl}${pluralName}`;
                }
            } catch {
                // try the next version candidate
            }
        }

        throw `Could not resolve Web API base url for table '${entityLogicalName}'.`;
    }

    async function getViewWebApiUrl(entityLogicalName, viewId, viewType, preferredVersion) {
        let queryParamName = '';

        if (viewType == 4230) {
            queryParamName = 'userQuery'
        } else if (viewType == 1039) {
            queryParamName = 'savedQuery';
        } else {
            throw 'unknown view type: ' + viewType;
        }

        const baseUrl = await getWebApiBaseUrl(entityLogicalName, preferredVersion);

        return `${baseUrl}?${queryParamName}=${viewId}`;;
    }

    async function getSingleRowApiUrl() {
        const entityLogicalName = Xrm.Page.data.entity.getEntityName();
        const versionArray = Xrm.Utility.getGlobalContext().getVersion().split('.');
        const version = versionArray[0] + '.' + versionArray[1];

        const baseUrl = await getWebApiBaseUrl(entityLogicalName, version);

        const recordId = Xrm.Page.data.entity.getId().replace('{', '').replace('}', '');

        return `${baseUrl}(${recordId})`;
    }

    async function getSingleRowApiUrlFromLocation() {
        const urlObj = new URL(window.location.href);
        const pageType = urlObj.searchParams.get('pagetype');
        const entityLogicalName = urlObj.searchParams.get('etn');
        const id = normalizeGuid(urlObj.searchParams.get('id'));

        if (pageType !== 'entityrecord' || !entityLogicalName || !id) {
            return null;
        }

        const baseUrl = await getWebApiBaseUrl(entityLogicalName);
        return `${baseUrl}(${id})`;
    }

    function getDataverseUrl() {
        const currentEnvironmentId = location.href.split('/environments/').pop().split('?')[0].split('/')[0];

        for (let i = 0; i < localStorage.length; i++) {
            const value = localStorage.getItem(localStorage.key(i));

            try {
                if (value.indexOf(currentEnvironmentId) === -1) {
                    continue
                }
                const valueJson = JSON.parse(value);
                if (Array.isArray(valueJson)) {
                    const environment = valueJson.filter(v => v.name === currentEnvironmentId)[0];
                    if (environment != null) {
                        return environment?.properties?.linkedEnvironmentMetadata?.instanceUrl;
                    }
                }
            } catch {
                // ignore
            }
        }
    }

    let urlToOpen = '';

    const urlObj = new URL(window.location.href);
    const viewId = urlObj.searchParams.get('viewid');
    const entityLogicalName = urlObj.searchParams.get('etn');
    const viewType = urlObj.searchParams.get('viewType');

    if (window.Xrm && window.Xrm.Page) {
        try {
            // check if on view
            if (viewId && entityLogicalName && viewType) {
                let preferredVersion = null;
                if (window.Xrm.Utility?.getGlobalContext) {
                    const versionArray = Xrm.Utility.getGlobalContext().getVersion().split('.');
                    preferredVersion = versionArray[0] + '.' + versionArray[1];
                }
                urlToOpen = await getViewWebApiUrl(entityLogicalName, viewId, viewType, preferredVersion);
            } else if (window.Xrm.Page.data?.entity && window.Xrm.Utility?.getGlobalContext) {
                urlToOpen = await getSingleRowApiUrl();
            } else {
                urlToOpen = await getSingleRowApiUrlFromLocation();
            }
        } catch (err) {
            alert(err);
            return;
        }

        if (urlToOpen) {
            urlToOpen += '#p'; // add the secret sauce
            window.postMessage({ action: 'openInWebApi', url: urlToOpen });
            return;
        }
    } else if (viewId && entityLogicalName && viewType) {
        try {
            urlToOpen = await getViewWebApiUrl(entityLogicalName, viewId, viewType);
            urlToOpen += '#p';
            window.postMessage({ action: 'openInWebApi', url: urlToOpen });
            return;
        } catch (err) {
            alert(err);
            return;
        }
    } else {
        try {
            const fastOpenRowUrl = await getSingleRowApiUrlFromLocation();
            if (fastOpenRowUrl) {
                window.postMessage({ action: 'openInWebApi', url: fastOpenRowUrl + '#p' });
                return;
            }
        } catch {
            // ignore and continue with regular fallbacks
        }
    }

    if (window.Xrm && window.Xrm.Page) {
        if (!urlToOpen) {
            alert(`Please open a form or view to use PrettifyMyWebApi`);
        }
    } else if (/\/api\/data\/v[0-9][0-9]?.[0-9]\//.test(window.location.pathname)) {
        // the host check is for supporting on-prem, where we always want to resort to the postmessage based flow
        // we only need total reload on the workflows table when viewing a single record, because of the monaco editor
        if (window.location.hash === '#p' && /\/api\/data\/v[0-9][0-9]?.[0-9]\/workflows\(/.test(window.location.pathname) && window.location.host.endsWith(".dynamics.com")) {
            window.location.reload();
        } else {
            window.location.hash = 'p';
            window.postMessage({ action: 'prettifyWebApi' });
        }
    }
    else if (window.location.host.endsWith('.powerautomate.com') || window.location.host.endsWith('.powerapps.com')) {
        const hrefToCheck = location.href + '/'; // append a slash in case the url ends with '/flows' or '/cloudflows'

        if (hrefToCheck.indexOf('flows/') === -1 || hrefToCheck.indexOf('/environments/') === -1) {
            return;
        }

        const instanceUrl = getDataverseUrl();

        if (!instanceUrl) {
            console.warn(`PrettifyMyWebApi: Couldn't find Dataverse instanceUrl.`);
            return;
        }

        // it can be /cloudflows/ or /flows/ so just check for flows/
        const flowUniqueId = hrefToCheck.split('flows/').pop().split('?')[0].split('/')[0];

        if (flowUniqueId && flowUniqueId.length === 36) {
            const url = instanceUrl + 'api/data/v9.2/workflows?$filter=resourceid eq ' + flowUniqueId + ' or workflowidunique eq ' + flowUniqueId + ' or workflowid eq ' + flowUniqueId + '#pf'
            window.postMessage({ action: 'openFlowInWebApi', url: url });
        } else {
            // for example, on make.powerapps it only works when viewing a flow from a solution
            console.warn(`PrettifyMyWebApi: Couldn't find Dataverse Flow Id.`);

            if (window.location.host.endsWith('make.powerapps.com') && hrefToCheck.indexOf('/solutions/') === -1) {
                alert(`Cannot find the Flow Id in the url. If you want to use this extension in make.powerapps.com, please open this Flow through a solution. Tip: if you use make.powerautomate.com, you should not run into this issue.`);
            }
        }
    }
})()
