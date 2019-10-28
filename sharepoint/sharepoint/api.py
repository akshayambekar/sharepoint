# -*- coding: utf-8 -*-
from datetime import datetime, timedelta
import logging
import os
from json import JSONDecodeError
from urllib.parse import urlsplit


def get_api_client(url, auth, logging=False):
    return APIclient(url, auth, logging)


def _remove_filename(path):
    """
    Return _dirname only if last element is a file. Don't include trailing slash in return value.
    :param path:
    :return:
    """
    parts = urlsplit(path)
    last = parts.path.split('/')[-1]
    if '.' in last:  # If last part of path has . then this is a file
        return os.path.dirname(path)
    else:
        if path[-1] == '/':
            path = path[:-1]
        return path


class APIclient(object):

    """
    SPApi is used to make API calls to the SharePoint server.

    Each SPObject includes a copy of SPApi and the api_url attribute of each copy is altered to point to the uri
    of the parent SPObject.  Copies are shallow so that only a single requests session exists in memory.
    Context management for requests session handled in SPSite object.


    """

    def __init__(self, url, auth, logging=False):

        # Connect to SharePoint
        self.auth = auth
        if self.auth.logged_in:
            self._session = self.auth.session  # If logged in, get requests session
        else:
            self._session = self.auth.login()  # auth.login returns requests session

        # Return JSON instead of XML
        self._session.headers.update({
            "Accept": "application/json; odata=verbose",
            "Content-type": "application/json; odata=verbose"
        })

        url = _remove_filename(url)
        self.api_url = self._get_api_url(url)

        # Form digest expiration. An unexpired digest token is required for posting
        # Setting to now guarantees that a new digest token is requested upon the first post
        self.expire = datetime.now()

        # Configure Logging
        self.logger = self._create_logger(logging)

    def get(self, url, params=None):
        """
        Call _http with GET
        """

        r = self.http(url, 'GET', params)
        return r

    def post(self, url, data=None):
        """
        Call _http with POST
        :param url:
        :param data:
        :return:
        """
        r = self.http(url, 'POST', data)
        return r

    def _get_api_url(self, url):
        """
        Determines sharepoint_url to SharePoint api by looking up contextinfo property.
        Preserves path to SharePoint sub-sites.

        :param url: URL to any resource in desired SharePoint site
        :return: String with api sharepoint_url
        """
        # Get rid of all path items after and including "_layouts"
        path_items = url.split('/')
        try:
            layouts_idx = path_items.index('_layouts')
            url = '/'.join(path_items[:layouts_idx])
        except ValueError:  # _layouts does not exist
            pass

        req_url = url + '/_api/contextinfo'
        r = self._session.post(req_url)
        self._check_response(r)
        site = r.json()['d']['GetContextWebInformation']['WebFullUrl']
        return "{}/_api".format(site)

    def _check_response(self, r):
        """
        Check if status code is ok.
        :param r: requests.Response object
        :return: None
        """
        # TODO Need to check if re-authentication is needed here?
        if not r.ok:
            try:
                err_msg = r.json()['error']['message']['value'],
            except (KeyError, JSONDecodeError) as _:
                err_msg = None
            err_str = '{code}: {reason}\n{message}\nRequest: {request}'.format(
                code=r.status_code,
                reason=r.reason,
                message=err_msg,
                request=r.request.method + ' ' + r.url
            )
            raise Exception(err_str)

    def _create_logger(self, logging_on):
        if logging_on:
            handler = logging.FileHandler('log')
        else:
            handler = logging.NullHandler()

        # create logger
        logger = logging.getLogger(__name__)
        logger.setLevel(logging.INFO)

        # create console handler and set level to debug
        handler.setLevel(logging.INFO)

        # create formatter
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

        # add formatter to ch
        handler.setFormatter(formatter)

        # add ch to logger
        logger.addHandler(handler)

        return logger

    def _digest(self):
        """
        If digest token is expired, obtain new value and update session header
        :return:
        """
        if self.expire <= datetime.now():
            r = self._session.post(self.api_url + '/contextinfo')
            self._check_response(r)
            contextinfo = r.json()['d']['GetContextWebInformation']
            self._session.headers['X-RequestDigest'] = contextinfo['FormDigestValue']
            self.expire = datetime.now() + timedelta(seconds=contextinfo['FormDigestTimeoutSeconds'])

    def http(self, url, verb, payload=None):
        """
        Make an http request
        :param url:
        :param verb:
        :return:
        """
        self.logger.debug("{0} request to {1}".format(verb, url))
        if verb.upper() == 'GET':
            r = self._session.get(url, params=payload)
        elif verb.upper() == 'POST':
            self._digest()
            r = self._session.post(url, data=payload)
        else:
            raise Exception('HTTP verb {} not supported'.format(verb))
        self._check_response(r)
        return r
