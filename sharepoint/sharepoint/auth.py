# -*- coding: utf-8 -*-
import re
import requests
from requests.exceptions import MissingSchema
from urllib.parse import urlsplit

from bs4 import BeautifulSoup

import pdb
# Suppress SSL verify warnings
try:
    from requests.packages.urllib3.exceptions import InsecureRequestWarning

    requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
except ModuleNotFoundError:
    pass


def manual_auth(getpass):
    """
    manual_auth is a factory function that chooses the correct Auth class
    based on the supplied sharepoint_url.

    """
    return ManualSPAuth(getpass)


class ManualSPAuth(object):
    """
    Manual authentication using user name and password.
    Credentials are securely received using the getpass package and transmitted
    by an https post.
    
    Jupyter uses it's own version of getpass. Since this special getpass
    function can't be accessed from the package level, it is instead passed
    into this function from Jupyter.
    """
    sharepoint_url = "https://fqdn.of.your.sharepoint"
    auth_failed_url = "https://fqdn.of.your.sharepoint/htdocs/public/auth_failed.html"

    def __init__(self, getpass):
        self.getpass = getpass
        self.logged_in = False
        self.session = requests.session()
        self.session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko)'
                                                   ' Chrome/61.0.3163.100 Safari/537.36'})
        self.session.verify = False  # Turn off SSL verification if using non-standard root CA

    def login(self):
        r = self._get_login_page()
        r = self._enter_credentials(r)  # post credentials with login form

        # Follow redirect
        r = self._submit_form(r)

        # Avoid MS 'keep me signed in' form
        r = self.session.get(self.sharepoint_url)

        # Follow redirect
        r = self._submit_form(r)

        print("Login successful")
        self.logged_in = True

        return self.session

    def _enter_credentials(self, r):

        # Enter credentials
        try:
            user = input('Enter username: ')
        except EOFError:
            raise Exception("Not inside interactive session. Can't get user input.")
        pw = self.getpass('Enter password: ')
        r = self._submit_form(r, {'USER': user, 'PASSWORD': pw})

        if r.url == self.auth_failed_url:
            raise Exception('Authentication Failed')
        return r

    def _get_login_page(self):

        # Request sharepoint_url. Server will redirect to auth page
        r = self.session.get(self.sharepoint_url)

        # Follow javascript redirect
        regex = re.compile('https://[^"]*')
        redirect = regex.findall(r.text)[1]
        r = self.session.get(redirect)
        return r

    def _submit_form(self, r, data=None):
        """
        Used for authentication. Get hidden form values and submit form. User values to form can be supplied
        with data argument.

        _session: requests session object
        r: response object with form
        data: dict of any user form data
        """

        # Add hidden values to payload
        post = re.compile('post', re.IGNORECASE)
        soup = BeautifulSoup(r.text, 'lxml')
        tags = soup.findAll('input', {'type': 'hidden'})
        payload = {tag.attrs['name']: tag.attrs['value'] for tag in tags}

        # Add user data to payload
        try:
            payload.update(data)
        except TypeError:  # No user supplied data
            pass

        # Get response sharepoint_url
        url = soup.findAll('form', method=post)[0].attrs['action']

        try:
            r = self.session.post(url, data=payload)
        except MissingSchema:  # If schema is missing, use domain from response sharepoint_url
            url_parts = urlsplit(r.url)
            url = '{url.scheme}://{url.netloc}{path}'.format(url=url_parts, path=url)
            r = self.session.post(url, data=payload)

        return r
