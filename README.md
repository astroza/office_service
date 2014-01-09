office_service
==============

Office Service is a server for dealing with Microsoft Office Word Files via a HTTP interface. Currently, it's able to convert a Word file to PDF and clean up personal information from Word file metadata. It use Microsoft Office COM API to access to an Word Application instance and do actions over the documents programmatically.

Motivation
------------

OpenOffice provides an API called UNO for manipulating documents programmatically. A OpenOffice functionality is convert Word files to PDF but it fails in rendering process. The result is a bad looking PDF that is not like the original one. My solution is to use the same program used to create and edit these documents.

Dependencies
------------

* Python (Lastest python 2.7 for windows http://python.org/download/)
* pip (http://stackoverflow.com/questions/4750806/how-to-install-pip-on-windows)
* Bottle (pip install bottle)
* Paste (pip install paste)
* pywin32 (http://sourceforge.net/projects/pywin32/files/pywin32/Build%20218/)
* A legal copy [ :-) ] of Microsoft Office 2007 (tested) or 2010 (untested)

How it works
------------
