# -*- coding: utf-8 -*-
import os
from urllib.parse import parse_qs, urlsplit
import uuid

from sharepoint import manual_auth
from sharepoint import get_api_client

import logging
logger = logging.getLogger(__name__)


def site_login(url, getpass, logging=False):
    auth = manual_auth(getpass)
    auth.login()
    api_client = get_api_client(url, auth, logging)
    site_url = api_client.api_url + '/Web'
    return SPSite(site_url, api_client)


def get_new_site(url, old_site, logging=False):
    """
    Reuse Authentication from old_site to create new site object
    """
    api_client = get_api_client(url, old_site._api_client.auth, logging)
    site_url = api_client.api_url + '/Web'
    return SPSite(site_url, api_client)


def _json_to_object(json, api_client):
    sp_type = json['__metadata']['type'].replace(".", "")
    logger.debug("Attribute is an object of type {}. Creating object.".format(sp_type))
    try:  # Check if class is implemented in python
        sp_class = globals()[sp_type]
    except KeyError:
        sp_class = SPObject
    return sp_class(json['__metadata']['uri'], api_client, json)


def _stringify(obj):
    try:  # If object is a string, encode apostrophe and place single quotes around it
        obj = obj.replace("'", "%27%27")  # Encoding apostrophe prevents interference early termination of quotes
        return "'" + obj + "'"
    except AttributeError:  # If object is not a string, convert it to a string without quotes
        if type(obj) == uuid.UUID:
            return "guid'{}'".format(str(obj))
        else:
            return str(obj)


class SPObject(object):

    @staticmethod
    def _make_call_string(method_name, *args, **kwargs):
        """
        Call a method from the SharePoint API.
        :param endpoint_url: URL for SharePoint resource with desired method. URI should not contain actual method.
        :param method_name: Name of method.
        :param args: Unnamed arguments for method.
        :param payload: OData query string for GET or data for POST
        :param verb: 'GET' or 'POST'
        :param kwargs: named arguments
        :return: Dict for json response
        """
        args = ', '.join([_stringify(arg) for arg in args])
        kwargs = ', '.join(["{0}={1}".format(key, _stringify(val)) for key, val in kwargs.items()])
        all_args = ', '.join([args, kwargs]).strip(', ')
        call_str = "/{name}({args})".format(
            name=method_name,
            args=all_args
        )
        return call_str

    def __init__(self, url, api_client, json=None):
        # json is dict with attribute values. Don't include 'd' key with json
        logger.debug('Creating new object from {}'.format(url))
        self._api_client = api_client
        self._endpoint_url = url
        self._attributes = {}
        if json:
            logger.debug('JSON was supplied. Parsing into attributes')
            json.pop('__metadata')
            for name, val in json.items():
                logger.debug('Setting attribute {}'.format(name))
                if type(val) is dict:
                    # Attribute is deferred so ignore it
                    pass
                else:
                    self._attributes[name] = val

    def __repr__(self):
        return self._repr_text(self._endpoint_url)

    def _repr_text(self, path_text):
        return "{0} {1}".format(type(self), path_text)

    def attribute(self, name):
        logger.debug("Retrieving attribute: {}".format(name))
        # Retrieve value if already stored
        if name in self._attributes.keys():
            logger.debug("Attribute was already stored. Retrieving value.")
            return self._attributes[name]
        else:
            url = self._attribute_url(name)
            logger.debug("Getting attribute from: {}".format(url))
            self._attributes[name] = LazyAttribute(url, self._api_client, name).value()
            return self._attributes[name]

    def lazy_attribute(self, name):
        url = self._attribute_url(name)
        logger.debug("Creating LazyAttribute for {0} from {1}".format(name, url))
        attribute = LazyAttribute(url, self._api_client, name)
        return attribute

    def _attribute_url(self, name):
        return self._endpoint_url + '/' + name

    def _method_get(self, method_name, *args, **kwargs):
        req_string = self._make_call_string(method_name, *args, **kwargs)
        url = self._endpoint_url + req_string
        return SPObject(url, self._api_client)

    def _method_post(self, method_name, *args, data=None, **kwargs):
        req_string = self._make_call_string(method_name, *args, **kwargs)
        url = self._endpoint_url + req_string
        return LazyPost(url, self._api_client, data)


class LazyAttribute(SPObject):

    def __init__(self, url, api_client, name, value=None):
        self._name = name
        self._value = self._parse_json(value)
        super(LazyAttribute, self).__init__(url, api_client)

    def value(self, query=None):

        if not self._value:
            json = self._api_client.get(self._endpoint_url, query).json()['d']
            logger.debug("Response for attribute {0}:\n{1}".format(self._name, json))
            self._value = self._parse_json(json)
        return self._value

    def _parse_json(self, json):
        logger.debug('Parsing json')
        if not json:
            logger.debug('Value is not json. Setting value to None')
            value = None
        elif self._name in json.keys():
            logger.debug('Attribute found in json.')
            value = json[self._name]
        elif 'results' in json.keys():  # json is a list
            results = [result for result in json['results']]
            logger.debug('Attribute is a list with length {0}. Parsing each item.'.format(len(json['results'])))
            value = [_json_to_object(item, self._api_client) for item in results]
        else:  # json represents object. Assign correct object class.
            value = _json_to_object(json, self._api_client)

        return value


class LazyPost(SPObject):

    def __init__(self, url, api_client, data=None):
        self.data = data
        super(LazyPost, self).__init__(url, api_client)

    def send(self):
        return self._api_client.post(self._endpoint_url, data=self.data)


class SPSite(SPObject):
    """
    SPSite represents the "Web" core endpoint from the SharePoint API.

    Users should only directly make instances of this class.
    All other class instances will be ultimately created through methods from this class.
    """

    def __repr__(self):
        return self._repr_text(self.attribute('Title'))

    def download_file(self, sp_file_path, destination='.'):
        """
        Convenience method that combines SPSite.get_file and SPFile.download file.
        """
        sp_file = self.get_file_by_path(sp_file_path)
        sp_file.download(destination)

    def get_file_by_path(self, file_path):
        file_path = self._append_site_path(file_path)
        file_url = self._method_get('GetFileByServerRelativeUrl', ServerRelativeUrl=file_path)._endpoint_url
        return SPFile(file_url, self._api_client)

    def get_file_by_id(self, file_id):
        file_url = self._method_get('GetFileById', uniqueId=file_id)._endpoint_url
        return SPFile(file_url, self._api_client)
    
    def get_file_by_url(self, url):
        """
        Attempt to parse sharepoint_url and retrieve file resource.
        If file_path or file id is known other methods should be used since this method is less robust.
        """
        parts = urlsplit(url)
        query = parse_qs(parts.query)
        keys = query.keys()
        if "sourcedoc" in keys:
            uid = query['sourcedoc'][0][1:-1]
            return self.get_file_by_id(uid)
        elif "SourceUrl" in keys:
            path = query['SourceUrl'][0] 
            path = '/' + '/'.join(path.split('/')[3:])
            # Check for invalid .xlsf extension
            base, ext = os.path.splitext(path)
            if ext == '.xlsf':
                path = base + '.xls'
            return self.get_file_by_path(path)
        else:  # Assume sharepoint_url is valid and remove all query items
            return self.get_file_by_path(parts.path)

    def get_folder(self, folder_path):
        folder_path = self._append_site_path(folder_path)
        new_url = self._method_get('GetFolderByServerRelativeUrl', folder_path)._endpoint_url
        return SPFolder(new_url, self._api_client)

    def _append_site_path(self, path):
        """

        :param path:
        :return:
        """
        if not path.startswith(self.attribute('ServerRelativeUrl')):
            path = self.attribute('ServerRelativeUrl') + '/' + path.strip('/')
        return path


class SPFolder(SPObject):

    def __repr__(self):
        return self._repr_text(self.attribute('ServerRelativeUrl'))

    def download_files(self, destination='.'):
        """
        Download all files in given folder. Ignores sub-folders.

        :param destination: Destination path where files will be saved.
        :return: None
        """
        for file in self.attribute('Files'):
            if os.path.splitext(file.attribute('Name'))[1] not in ['.aspx']:  # Downloading .aspx results in 403 forbidden error
                file.download(destination)

    def download(self, destination='.', maxdepth=None):
        """
        Download all files and sub-folders up to specified depth.
        :param destination: Top-level path where files will be saved.
        :param maxdepth: Number of sub-folder levels to retrieve. To get all subfolders, set level to -1.
        :return: None
        """

        base_path = os.path.dirname(self.attribute('ServerRelativeUrl'))

        destination = destination.strip('/')
        logger.debug("Start download of {}".format(base_path))
        for folder, _, _, in self.walk(maxdepth=maxdepth):
            logger.debug('Inside {}'.format(folder.attribute('Name')))
            folder_path = folder.attribute('ServerRelativeUrl')[len(base_path):]
            logger.debug('Folder path: {}'.format(folder_path))
            dest_folder = destination + '/' + folder_path
            folder.download_files(dest_folder)

    def listdir(self):
        return self.attribute('Files') + self.attribute('Folders')

    def walk(self, topdown=True, maxdepth=None):
        """
        Analogous to os.walk
        :param topdown:
        :param maxdepth: Maximum recursion depth.
        :return:
        """
        top = self
        folders = self.attribute('Folders')
        logger.debug("Inside walk. Folders retrieved.")
        files = self.attribute('Files')
        logger.debug("Inside walk. Files retrieved.")

        if topdown:
            logger.debug("Reached walk endpoint.")
            yield top, folders, files

        if maxdepth is None or maxdepth > 1:
            for folder in folders:
                if maxdepth:
                    newdepth = maxdepth - 1
                else:
                    newdepth = None
                for x in folder.walk(topdown, newdepth):
                    yield x
        if not topdown:
            yield top, folders, files

    def upload_file(self, filename, overwrite=True):
        """
        Upload a file to SharePoint
        :param filename: String with path to local file
        :param overwrite: Overwrite existing file on SharePoint?
        :return: None
        """
        file_size = os.path.getsize(filename)
        chunk_size = 1024*1024

        file_base = os.path.split(filename)[1]
        stream = False
        if file_size <= chunk_size:
            with open(filename, 'rb') as f:
                file = f.read()
        else:
            file = None  # Don't include data with add method. Send it via streaming.
            stream = True

        try:
            # This runs even if streaming upload is used because an empty file must be created before streaming starts
            r = self.lazy_attribute('Files')._method_post('add',
                                                          data=file,
                                                          url=file_base,
                                                          overwrite=str(overwrite).lower(),
                                                          ).send()
            file = _json_to_object(r.json()['d'], self._api_client)

        except:
            logger.exception("Upload Failed")
            return
            
        if stream:
            self._stream_upload(filename, file_size, chunk_size)

        print("Uploaded {0} to {1}".format(filename, self.attribute('ServerRelativeUrl')))
        if file.attribute('CheckOutType') != 2:
            logger.debug("File {} is checked out. Checking in file.".format(file.attribute('Name')))
            file._method_post('CheckIn', comment="", checkInType=0).send()  # Check in type 0 is minor check in



    def _stream_upload(self, filename, file_size, chunk_size):
        """
        For larger files, upload in chunks.
        :param filename:
        :param file_size:
        :param chunk_size:
        :return:
        """
        guid = uuid.uuid4()
        first_chunk = True
        f = open(filename, 'rb')
        i = self._endpoint_url.find("Web") + 3
        site_url = self._endpoint_url[:i]
        # Add empty file to folder
        relative_path = self.attribute('ServerRelativeUrl')
        file = (SPObject(site_url, self._api_client).  # create site level object to access GetFileBy... method
                _method_get('GetFileByServerRelativeUrl', ServerRelativeUrl=relative_path + '/' + filename))
        try:
            offset = 0
            while True:
                data = f.read(chunk_size)
                if first_chunk:
                    file._method_post('startupload', uploadId=guid, data=data).send()
                    first_chunk = False
                elif offset >= file_size - chunk_size:
                    file._method_post('finishupload', uploadId=guid, fileOffset=offset, data=data).send()
                    break
                else:
                    file._method_post('continueupload', uploadId=guid, fileOffset=offset, data=data).send()
                offset += len(data)
        except:
            f.close()
            logger.debug('Upload Failed')
            file._method_post('cancelupload', guid).send()


class SPFile(SPObject):

    def __repr__(self):
        return self._repr_text(self.attribute('ServerRelativeUrl'))

    def download(self, destination='.'):
        """
        Download single file from SharePoint.

        :param destination: String with path where file will be saved.
        :return: None
        """

        # Download file
        # api_client is directly used instead of attribute because $value returns raw data rather than json
        logger.debug("Starting download: {}".format(self.attribute('Name')))
        r = self._api_client.get(self._endpoint_url + '/$value')
        logger.debug("Download complete")
        destination = destination.strip('/')
        destination = os.path.abspath(destination)

        if not os.path.isdir(destination):
            os.makedirs(destination)

        destination = os.path.join(destination, self.attribute('Name'))
        logger.debug("Writing file to disk")
        with open(destination, 'wb') as f:
            for chunk in r.iter_content(chunk_size=128):
                f.write(chunk)
        
        print("Successfully downloaded file as {0}".format(destination))