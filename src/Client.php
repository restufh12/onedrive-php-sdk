<?php

namespace Krizalys\Onedrive;

use GuzzleHttp\ClientInterface;
use Krizalys\Onedrive\Proxy\DriveItemProxy;
use Krizalys\Onedrive\Proxy\DriveProxy;
use Microsoft\Graph\Graph;
use Microsoft\Graph\Model;

/**
 * @class Client
 *
 * A Client instance allows communication with the OneDrive API and perform
 * operations programmatically.
 *
 * To manage your Live Connect applications, see here:
 * https://apps.dev.microsoft.com/#/appList
 */
class Client
{
    /**
     * @var string
     *      The base URL for authorization requests.
     */
    const AUTH_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize';

    /**
     * @var string
     *      The base URL for token requests.
     */
    const TOKEN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';

    /**
     * @var string
     *      The client ID.
     */
    private $clientId;

    /**
     * @var Microsoft\Graph\Graph
     *      The Microsoft Graph.
     */
    private $graph;

    /**
     * @var GuzzleHttp\ClientInterface
     *      The Guzzle HTTP client.
     */
    private $httpClient;

    /**
     * @var object
     *      The OAuth state (token, etc...).
     */
    private $_state;

    /**
     * Constructor.
     *
     * @param string $clientId
     *        The client ID.
     * @param Microsoft\Graph\Graph $graph
     *        The graph.
     * @param GuzzleHttp\ClientInterface $httpClient
     *        The HTTP client.
     * @param array $options
     *        The options to use while creating this object.
     *        Valid supported keys are:
     *          - 'state' (object) When defined, it should contain a valid
     *            OneDrive client state, as returned by getState(). Default: [].
     *
     * @throws Exception
     *         Thrown if $clientId is null.
     */
    public function __construct(
        $clientId,
        Graph $graph,
        ClientInterface $httpClient,
        array $options = []
    ) {
        if ($clientId === null) {
            throw new \Exception('The client ID must be set');
        }

        $this->clientId   = $clientId;
        $this->graph      = $graph;
        $this->httpClient = $httpClient;

        $this->_state = array_key_exists('state', $options)
            ? $options['state'] : (object) [
                'redirect_uri' => null,
                'token'        => null,
            ];
    }

    /**
     * Gets the current state of this Client instance. Typically saved in the
     * session and passed back to the Client constructor for further requests.
     *
     * @return object
     *         The state of this Client instance.
     */
    public function getState()
    {
        return $this->_state;
    }

    /**
     * Gets the URL of the log in form. After login, the browser is redirected
     * to the redirect URI, and a code is passed as a query string parameter to
     * this URI.
     *
     * The browser is also redirected to the redirect URI if the user is already
     * logged in.
     *
     * @param array $scopes
     *        The OneDrive scopes requested by the application. Supported
     *        values:
     *          - 'offline_access'
     *          - 'files.read'
     *          - 'files.read.all'
     *          - 'files.readwrite'
     *          - 'files.readwrite.all'
     * @param string $redirectUri
     *        The URI to which to redirect to upon successful log in.
     *
     * @return string
     *         The log in URL.
     */
    public function getLogInUrl(array $scopes, $redirectUri)
    {
        $redirectUri                = (string) $redirectUri;
        $this->_state->redirect_uri = $redirectUri;

        $values = [
            'client_id'     => $this->clientId,
            'response_type' => 'code',
            'redirect_uri'  => $redirectUri,
            'scope'         => implode(' ', $scopes),
            'response_mode' => 'query',
        ];

        $query = http_build_query($values, '', '&', PHP_QUERY_RFC3986);

        // When visiting this URL and authenticating successfully, the agent is
        // redirected to the redirect URI, with a code passed in the query
        // string (the name of the variable is "code"). This is suitable for
        // PHP.
        return self::AUTH_URL . "?$query";
    }

    /**
     * Gets the access token expiration delay.
     *
     * @return int
     *         The token expiration delay, in seconds.
     */
    public function getTokenExpire()
    {
        return $this->_state->token->obtained
            + $this->_state->token->data->expires_in - time();
    }

    /**
     * Gets the status of the current access token.
     *
     * @return int
     *         The status of the current access token:
     *           -  0 No access token.
     *           - -1 Access token will expire soon (1 minute or less).
     *           - -2 Access token is expired.
     *           -  1 Access token is valid.
     */
    public function getAccessTokenStatus()
    {
        if (null === $this->_state->token) {
            return 0;
        }

        $remaining = $this->getTokenExpire();

        if (0 >= $remaining) {
            return -2;
        }

        if (60 >= $remaining) {
            return -1;
        }

        return 1;
    }

    /**
     * Obtains a new access token from OAuth. This token is valid for one hour.
     *
     * @param string $clientSecret
     *        The OneDrive client secret.
     * @param string $code
     *        The code returned by OneDrive after successful log in.
     *
     * @throws Exception
     *         Thrown if the redirect URI of this Client instance's state is not
     *         set.
     * @throws Exception
     *         Thrown if the HTTP response body cannot be JSON-decoded.
     */
    public function obtainAccessToken($clientSecret, $code)
    {
        if (null === $this->_state->redirect_uri) {
            throw new \Exception(
                'The state\'s redirect URI must be set to call'
                    . ' obtainAccessToken()'
            );
        }

        $values = [
            'client_id'     => $this->clientId,
            'redirect_uri'  => $this->_state->redirect_uri,
            'client_secret' => (string) $clientSecret,
            'code'          => (string) $code,
            'grant_type'    => 'authorization_code',
        ];

        $response = $this->httpClient->post(
            self::TOKEN_URL,
            ['form_params' => $values]
        );

        $body = $response->getBody();
        $data = json_decode($body);

        if ($data === null) {
            throw new \Exception('json_decode() failed');
        }

        $this->_state->redirect_uri = null;

        $this->_state->token = (object) [
            'obtained' => time(),
            'data'     => $data,
        ];

        $this->graph->setAccessToken($this->_state->token->data->access_token);
    }

    /**
     * Renews the access token from OAuth. This token is valid for one hour.
     *
     * @param string $clientSecret
     *        The client secret.
     */
    public function renewAccessToken($clientSecret)
    {
        if (null === $this->_state->token->data->refresh_token) {
            throw new \Exception(
                'The refresh token is not set or no permission for'
                    . ' \'wl.offline_access\' was given to renew the token'
            );
        }

        $values = [
            'client_id'     => $this->clientId,
            'client_secret' => $clientSecret,
            'grant_type'    => 'refresh_token',
            'refresh_token' => $this->_state->token->data->refresh_token,
        ];

        $response = $this->httpClient->post(
            self::TOKEN_URL,
            ['form_params' => $values]
        );

        $body = $response->getBody();
        $data = json_decode($body);

        if ($data === null) {
            throw new \Exception('json_decode() failed');
        }

        $this->_state->token = (object) [
            'obtained' => time(),
            'data'     => $data,
        ];
    }

    /**
     * @return array
     *         The drives.
     */
    public function getDrives()
    {
        $driveLocator = '/me/drives';
        $endpoint     = "$driveLocator";

        $response = $this
            ->graph
            ->createCollectionRequest('GET', $endpoint)
            ->execute();

        $status = $response->getStatus();

        if ($status != 200) {
            throw new \Exception("Unexpected status code produced by 'GET $endpoint': $status");
        }

        $drives = $response->getResponseAsObject(Model\Drive::class);

        if (!is_array($drives)) {
            return [];
        }

        return array_map(function (Model\Drive $drive) {
            return new DriveProxy($this->graph, $drive);
        }, $drives);
    }

    /**
     * @return DriveProxy
     *         The drive.
     */
    public function getMyDrive()
    {
        $driveLocator = '/me/drive';
        $endpoint     = "$driveLocator";

        $response = $this
            ->graph
            ->createRequest('GET', $endpoint)
            ->execute();

        $status = $response->getStatus();

        if ($status != 200) {
            throw new \Exception();
        }

        $drive = $response->getResponseAsObject(Model\Drive::class);

        return new DriveProxy($this->graph, $drive);
    }

    /**
     * @param string $driveId
     *        The drive ID.
     *
     * @return DriveProxy
     *         The drive.
     */
    public function getDriveById($driveId)
    {
        $driveLocator = "/drives/$driveId";
        $endpoint     = "$driveLocator";

        $response = $this
            ->graph
            ->createRequest('GET', $endpoint)
            ->execute();

        $status = $response->getStatus();

        if ($status != 200) {
            throw new \Exception();
        }

        $drive = $response->getResponseAsObject(Model\Drive::class);

        return new DriveProxy($this->graph, $drive);
    }

    /**
     * @param string $idOrUserPrincipalName
     *        The ID or user principal name.
     *
     * @return DriveProxy
     *         The drive.
     */
    public function getDriveByUser($idOrUserPrincipalName)
    {
        $userLocator  = "/users/$idOrUserPrincipalName";
        $driveLocator = '/drive';
        $endpoint     = "$userLocator$driveLocator";

        $response = $this
            ->graph
            ->createRequest('GET', $endpoint)
            ->execute();

        $status = $response->getStatus();

        if ($status != 200) {
            throw new \Exception();
        }

        $drive = $response->getResponseAsObject(Model\Drive::class);

        return new DriveProxy($this->graph, $drive);
    }

    /**
     * @param string $groupId
     *        The group ID.
     *
     * @return DriveProxy
     *         The drive.
     */
    public function getDriveByGroup($groupId)
    {
        $groupLocator = "/groups/$groupId";
        $driveLocator = '/drive';
        $endpoint     = "$groupLocator$driveLocator";

        $response = $this
            ->graph
            ->createRequest('GET', $endpoint)
            ->execute();

        $status = $response->getStatus();

        if ($status != 200) {
            throw new \Exception();
        }

        $drive = $response->getResponseAsObject(Model\Drive::class);

        return new DriveProxy($this->graph, $drive);
    }

    /**
     * @param string $siteId
     *        The site ID.
     *
     * @return DriveProxy
     *         The drive.
     */
    public function getDriveBySite($siteId)
    {
        $siteLocator  = "/sites/$siteId";
        $driveLocator = '/drive';
        $endpoint     = "$siteLocator$driveLocator";

        $response = $this
            ->graph
            ->createRequest('GET', $endpoint)
            ->execute();

        $status = $response->getStatus();

        if ($status != 200) {
            throw new \Exception();
        }

        $drive = $response->getResponseAsObject(Model\Drive::class);

        return new DriveProxy($this->graph, $drive);
    }

    /**
     * @param string $driveId
     *        The drive ID.
     * @param string $itemId
     *        The drive item ID.
     *
     * @return DriveItemProxy
     *         The drive item.
     */
    public function getDriveItemById($driveId, $itemId)
    {
        $driveLocator = "/drives/$driveId";
        $itemLocator  = "/items/$itemId";
        $endpoint     = "$driveLocator$itemLocator";

        $response = $this
            ->graph
            ->createRequest('GET', $endpoint)
            ->execute();

        $status = $response->getStatus();

        if ($status != 200) {
            throw new \Exception();
        }

        $driveItem = $response->getResponseAsObject(Model\DriveItem::class);

        return new DriveItemProxy($this->graph, $driveItem);
    }

    /**
     * @return DriveItemProxy
     *         The root drive item.
     */
    public function getRoot()
    {
        $driveLocator = '/me/drive';
        $itemLocator  = '/root';
        $endpoint     = "$driveLocator$itemLocator";

        $response = $this
            ->graph
            ->createRequest('GET', $endpoint)
            ->execute();

        $status = $response->getStatus();

        if ($status != 200) {
            throw new \Exception("Unexpected status code produced by 'GET $endpoint': $status");
        }

        $driveItem = $response->getResponseAsObject(Model\DriveItem::class);

        return new DriveItemProxy($this->graph, $driveItem);
    }

    /**
     * @param string $specialFolderName
     *        The special folder name. Supported values:
     *          - 'documents'
     *          - 'photos'
     *          - 'cameraroll'
     *          - 'approot'
     *          - 'music'
     *
     * @return DriveItemProxy
     *         The root drive item.
     */
    public function getSpecialFolder($specialFolderName)
    {
        $driveLocator         = '/me/drive';
        $specialFolderLocator = "/special/$specialFolderName";
        $endpoint             = "$driveLocator$specialFolderLocator";

        $response = $this
            ->graph
            ->createRequest('GET', $endpoint)
            ->execute();

        $status = $response->getStatus();

        if ($status != 200) {
            throw new \Exception("Unexpected status code produced by 'GET $endpoint': $status");
        }

        $driveItem = $response->getResponseAsObject(Model\DriveItem::class);

        return new DriveItemProxy($this->graph, $driveItem);
    }

    /**
     * @return array
     *         The shared drive items.
     */
    public function getShared()
    {
        $driveLocator = '/me/drive';
        $endpoint     = "$driveLocator/sharedWithMe";

        $response = $this
            ->graph
            ->createCollectionRequest('GET', $endpoint)
            ->execute();

        $status = $response->getStatus();

        if ($status != 200) {
            throw new \Exception("Unexpected status code produced by 'GET $endpoint': $status");
        }

        $driveItems = $response->getResponseAsObject(Model\DriveItem::class);

        if (!is_array($driveItems)) {
            return [];
        }

        return array_map(function (Model\DriveItem $driveItem) {
            return new DriveItemProxy($this->graph, $driveItem);
        }, $driveItems);
    }

    /**
     * @return array
     *         The recent drive items.
     */
    public function getRecent()
    {
        $driveLocator = '/me/drive';
        $endpoint     = "$driveLocator/recent";

        $response = $this
            ->graph
            ->createCollectionRequest('GET', $endpoint)
            ->execute();

        $status = $response->getStatus();

        if ($status != 200) {
            throw new \Exception("Unexpected status code produced by 'GET $endpoint': $status");
        }

        $driveItems = $response->getResponseAsObject(Model\DriveItem::class);

        if (!is_array($driveItems)) {
            return [];
        }

        return array_map(function (Model\DriveItem $driveItem) {
            return new DriveItemProxy($this->graph, $driveItem);
        }, $driveItems);
    }
}
