office_service
==============

Office Service is a server for dealing with Microsoft Office Word files via a HTTP interface. Currently, it's able to convert a Word file to PDF and clean up personal information from Word file metadata. It use Microsoft Office COM API to access to an Word Application instance and do actions over the documents programmatically.

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

### HTTP Interface
There are two URL:
###### /to/pdf 
It receives a word file (.doc or .docx) via POST method and send a PDF file as response. Example: convert.html
###### /cleanup/word
It receives a word file (.doc or .docx) via POST method send another word file without personal information (but identical to the first one). Example: cleanup.html

Before deliver the response to client, the work is sent to the Work Queue to process the request upon Microsoft Word.

### Work Queue
Microsoft Word attends one request at a time. It was necessary to create a queue to deliver one work at a time. Under demand, the Word instance is kept alive and serves each request enqueued (get best performance). When the queue is empty the Word instance is destroyed for cleaning (get best stability)
