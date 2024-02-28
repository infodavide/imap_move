#!/usr/bin/python
# -*- coding: utf-8-
"""
Move message from one folder to another one in a targeted mailbox
"""
import argparse
import atexit
import signal
import imaplib
import logging
import os
import pathlib
import sys
import xml.etree.ElementTree as etree
from logging.handlers import RotatingFileHandler

__author__ = 'David Rolland, contact@infodavid.org'
__copyright__ = 'Copyright © 2023 David Rolland'
__license__ = 'MIT'

IMAP4_PORT: int = 143


class _ObjectView:
    """
    Wrapper of the object
    """

    def __init__(self, d):
        """
        Initialize
        :param d: the data
        """
        self.__dict__ = d

    def __str__(self) -> str:
        """ Returns the string representation of the view """
        return str(self.__dict__)


class ImapSettings:
    """
    IMAP Settings.
    """
    server: str = None  # Full name or IP address of your IMAP server
    use_ssl: bool = False  # Set True to use SSL for the IMAP server
    port: int = IMAP4_PORT  # Port of your IMAP server
    user: str = None  # User used to connect to your IMAP server
    password: str = None  # Password (base64 encoded) of the user used to connect to your IMAP server
    folder: str = None  # The folder from where to cut messages from or to paste messages into
    trash: str = None  # The trash folder of your IMAP server

    def parse(self, node: etree.Element, accounts: {}) -> None:
        """
        Parse XML
        :param node: the node
        :param accounts: the accounts
        """
        self.server = node.get('server')
        v = node.get('port')
        if v is not None:
            self.port = int(v)
        else:
            self.port = 143
        v = node.get('folder')
        if v is not None:
            self.folder = str(v)
        else:
            self.folder = '"[Gmail]/Sent Mail"'
        v = node.get('trash')
        if v is not None:
            self.trash = str(v)
        else:
            self.trash = '"[Gmail]/Trash"'
        self.use_ssl = node.get('ssl') == 'True' or node.get('ssl') == 'true'
        account_id: str = node.get('account-id')
        account = accounts[account_id]
        if account:
            self.user = account[0]
            self.password = account[1]


class Settings:
    """
    Settings used by the IMAP deletion.
    """
    source_server: ImapSettings = None  # settings of your source IMAP server
    target_server: ImapSettings = None  # settings of your target IMAP server
    path: str = None  # Path for the files used by the application
    log_path: str  # Path to the logs file, not used in this version
    log_level: str  # Level of logs, not used in this version

    def parse(self, path: str) -> None:
        """
        Parse the XML configuration.
        """
        with open(path, encoding='utf-8') as f:
            tree = etree.parse(f)
        root_node: etree.Element = tree.getroot()
        log_node: etree.Element = root_node.find('log')
        if log_node is not None:
            v = log_node.get('path')
            if v is not None:
                self.log_path = str(v)
            v = log_node.get('level')
            if v is not None:
                self.log_level = str(v)
        accounts = {}
        for node in tree.findall('accounts/account'):
            v1 = node.get('user')
            v2 = node.get('password')
            v3 = node.get('id')
            if v1 is not None and v2 is not None and v3 is not None:
                accounts[v3] = [v1, v2]
        imap_node: etree.Element = root_node.find('source')
        if imap_node is not None:
            self.source_server = ImapSettings()
            self.source_server.parse(imap_node, accounts)
        else:
            raise IOError('No source imap element specified in the XML configuration, refer to the XML schema')
        imap_node = root_node.find('target')
        if imap_node is not None:
            self.target_server = ImapSettings()
            self.target_server.parse(imap_node, accounts)
        else:
            raise IOError('No target imap element specified in the XML configuration, refer to the XML schema')
        self.path = os.path.dirname(path)


def create_rotating_log(path: str, level: str) -> logging.Logger:
    """
    Create the logger with file rotation
    :param path: the path of the main log file
    :param level: the log level as defined in logging module
    :return: the logger
    """
    result: logging.Logger = logging.getLogger("imap_move")
    path_obj: pathlib.Path = pathlib.Path(path)
    if not os.path.exists(path_obj.parent.absolute()):
        os.makedirs(path_obj.parent.absolute())
    if os.path.exists(path):
        with open(path, 'w', encoding='utf-8') as f:
            f.close()
    else:
        path_obj.touch()
    # noinspection Spellchecker
    formatter: logging.Formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    console_handler: logging.Handler = logging.StreamHandler()
    console_handler.setLevel(level)
    console_handler.setFormatter(formatter)
    result.addHandler(console_handler)
    file_handler: logging.Handler = RotatingFileHandler(path, maxBytes=1024 * 1024 * 5, backupCount=5)
    # noinspection PyUnresolvedReferences
    file_handler.setLevel(level)
    file_handler.setFormatter(formatter)
    result.addHandler(file_handler)
    # noinspection PyUnresolvedReferences
    result.setLevel(level)
    return result


def cleanup() -> None:
    """
    Cleanup the instances and session
    """
    logger.log(logging.INFO, "Cleaning...")
    if 'source_mailbox' in globals():
        if 'logger' in globals():
            logger.info('Source IMAP session state: %s', source_mailbox.state)
        if source_mailbox.state == 'SELECTED':
            logger.log(logging.DEBUG, 'Closing...')
            source_mailbox.expunge()
            source_mailbox.close()
            source_mailbox.logout()
    if 'target_mailbox' in globals():
        if 'logger' in globals():
            logger.info('Target IMAP session state: %s', target_mailbox.state)
        if target_mailbox.state == 'SELECTED':
            logger.log(logging.DEBUG, 'Closing...')
            target_mailbox.expunge()
            target_mailbox.close()
            target_mailbox.logout()


# pylint: disable=missing-type-doc
def signal_handler(sig=None, frame=None) -> None:
    """
    Trigger the cleanup when program is exited
    :param sig: the signal
    :param frame: the frame
    """
    cleanup()
# pylint: enable=missing-type-doc


parser = argparse.ArgumentParser(prog='imap_move.py', description='Move messages from IMAP server to another one')
parser.add_argument('-f', required=True, help='Configuration file')
parser.add_argument('-l', help='Log level', default='INFO')
parser.add_argument('-v', default=False, action='store_true', help='Verbose')
args = parser.parse_args()
LOG_LEVEL: str = args.l
if LOG_LEVEL.startswith('"') and LOG_LEVEL.endswith('"'):
    LOG_LEVEL = LOG_LEVEL[1:-1]
if LOG_LEVEL.startswith("'") and LOG_LEVEL.endswith("'"):
    LOG_LEVEL = LOG_LEVEL[1:-1]
CONFIG_PATH: str = args.f
if CONFIG_PATH.startswith('"') and CONFIG_PATH.endswith('"'):
    CONFIG_PATH = CONFIG_PATH[1:-1]
if CONFIG_PATH.startswith("'") and CONFIG_PATH.endswith("'"):
    CONFIG_PATH = CONFIG_PATH[1:-1]
if not os.path.exists(CONFIG_PATH):
    CONFIG_PATH = str(pathlib.Path(__file__).parent) + os.sep + CONFIG_PATH
LOG_PATH: str = os.path.splitext(CONFIG_PATH)[0] + '.log'
settings: Settings = Settings()
settings.log_path = LOG_PATH
settings.log_level = LOG_LEVEL
settings.parse(os.path.abspath(CONFIG_PATH))
logger = create_rotating_log(settings.log_path, settings.log_level)
logger.info('Using arguments: %s', repr(args))

if not args.f or not os.path.isfile(args.f):
    print('Input file is required and must be valid.')
    sys.exit(1)

LOCK_PATH: str = os.path.abspath(os.path.dirname(CONFIG_PATH)) + os.sep + '.imap_move.lck'
logger.info('Log level set to: %s', logging.getLevelName(logger.level))
atexit.register(signal_handler)
signal.signal(signal.SIGINT, signal_handler)
logger.info('Connecting to source server: %s:%s with user: %s', settings.source_server.server, str(settings.source_server.port), settings.source_server.user)

if settings.source_server.use_ssl:
    source_mailbox = imaplib.IMAP4_SSL(host=settings.source_server.server, port=settings.source_server.port)
else:
    source_mailbox = imaplib.IMAP4(host=settings.source_server.server, port=settings.source_server.port)

source_mailbox.login(settings.source_server.user, settings.source_server.password)

if logger.isEnabledFor(logging.DEBUG):
    buffer: str = 'Available folders on source:\n'
    for i in source_mailbox.list()[1]:
        p = i.decode().split(' "/" ')
        buffer += (p[0] + " = " + p[1]) + '\n'
    logger.log(logging.DEBUG, buffer)

logger.info('Connecting to target server: %s:%s with user: %s', settings.target_server.server, str(settings.target_server.port), settings.target_server.user)

if settings.target_server.use_ssl:
    target_mailbox = imaplib.IMAP4_SSL(host=settings.target_server.server, port=settings.target_server.port)
else:
    target_mailbox = imaplib.IMAP4(host=settings.target_server.server, port=settings.target_server.port)

target_mailbox.login(settings.target_server.user, settings.target_server.password)

if logger.isEnabledFor(logging.DEBUG):
    buffer: str = 'Available folders on target:\n'
    for i in target_mailbox.list()[1]:
        p = i.decode().split(' "/" ')
        buffer += (p[0] + " = " + p[1]) + '\n'
    logger.log(logging.DEBUG, buffer)
logger.info('Selecting folder on source: %s', settings.source_server.folder)
source_mailbox.select(settings.source_server.folder)
logger.info('Selecting folder on target: %s', settings.target_server.folder)
target_mailbox.select(settings.target_server.folder)
typ, data = source_mailbox.search(None, 'ALL')
count: int = 0

for num in data[0].split():
    logger.info('Fetching message: %s', str(num))
    resp, data = source_mailbox.fetch(num, "(FLAGS INTERNALDATE BODY.PEEK[])")
    message = data[0][1]
    logger.log(logging.DEBUG, 'Retrieving flags')
    flags = []
    for flag in imaplib.ParseFlags(data[0][0]):
        flags.append(flag.decode())
    flag_str = ' '.join(flags)
    logger.debug('Retrieving internal date')
    date = imaplib.Time2Internaldate(imaplib.Internaldate2tuple(data[0][0]))
    logger.debug('Moving message to folder: %s', settings.target_server.folder)
    append_result = target_mailbox.append(settings.target_server.folder, flag_str, date, message)

    if append_result and len(append_result) > 0 and str(append_result[0]).upper() == 'OK':
        count = count + 1
        source_mailbox.store(num, '+FLAGS', '\\Deleted')

logger.info('%s messages moved', str(count))
source_mailbox.select(settings.source_server.trash)  # select trash
source_mailbox.store("1:*", '+FLAGS', '\\Deleted')  # flag all trash as Deleted
sys.exit(0)
