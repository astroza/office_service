# (c) 2014 - Felipe Astroza Araya
# Under BSD License

import os
import win32com.client as win32
import pythoncom
import md5
import time
import threading
from bottle import route, run, request, static_file, Bottle, HTTPError

app = Bottle()
UPLOAD_PATH = os.getcwd()+'\\uploads'
PDF_PATH = os.getcwd()+'\\pdfs'

gc_list = []
GC_SLEEP_TIME = 60*5

class GCFile:
	def __init__(self, path):
		self.file_path = path
		self.creation_time = time.time()
		gc_list.append(self)

	def is_waste(self):
		if time.time() - self.creation_time >= GC_SLEEP_TIME:
			return True
		return False

def gc_thread():
	while(True):
		for f in gc_list:
			if f.is_waste():
				print "\nGC: Removing " + f.file_path + "\n"
				os.remove(f.file_path)
				gc_list.remove(f)
		time.sleep(GC_SLEEP_TIME)

class WorkQueueThread(threading.Thread):
	def __init__(self):
		self.queue_sem = threading.Semaphore(0)
		self.queue = []
		threading.Thread.__init__(self)

	def append(self, item):
		self.queue.append(item)
		self.queue_sem.release()

	def doc_to_pdf(self, item):
		doc = self.wapp.Documents.Open(item.word_file_path)
		pdf_path = PDF_PATH+'\\'+item.word_filename+'.pdf'
		doc.SaveAs(pdf_path, FileFormat=17)
		doc.Close(False)
		GCFile(pdf_path)

	def cleanup_doc(self, item):
		doc = self.wapp.Documents.Open(item.word_file_path)
		doc.RemoveDocumentInformation(4)
		doc.Save()
		doc.Close(False)

	def run(self):
		pythoncom.CoInitialize()
		self.wapp = None
		services = [self.doc_to_pdf, self.cleanup_doc]
		while(True):
			if self.queue_sem.acquire(False) == False:
				if self.wapp != None:
					self.wapp.Application.Quit()
					print "\nWorkQueueThread: Word Application was closed\n"
				self.queue_sem.acquire() # Espera por un trabajo
				self.wapp = win32.Dispatch('Word.Application')
				self.wapp.Visible = False
				self.wapp.DisplayAlerts = False
				print "\nWorkQueueThread: Work Application just started\n"

			item = self.queue.pop(0)
			print "\nWorkQueueThread: Processing %s service_type=%d\n" % (item.word_file_path, item.service_type)
			if item.service_type < len(services):
				services[item.service_type](item)
			else:
				print "\nWorkQueueThread: Invalid service_type(=%d)\n" % (item.service_type)
			item.lock.release()

class WorkQueueItem:
	def __init__(self, path, word_filename, orig_filename, filename_ext, service_type):
		self.word_file_path = path
		self.orig_filename = orig_filename
		self.filename_ext = filename_ext
		self.word_filename = word_filename
		self.service_type = service_type
		self.lock = threading.Semaphore(0)
		worker_thread.append(self)

	def wait_work(self):
		self.lock.acquire()

def save_word_file():
	upload = request.files.get('upload')
	name, ext = os.path.splitext(upload.filename)
	if ext not in ('.doc','.docx'):
		raise HTTPError(404, "Solo archivos MS Word son permitidos: " + ext)
	name_hash = md5.md5(name+str(time.time())).hexdigest()
	save_file = open(UPLOAD_PATH+'\\'+name_hash+ext, "wb")
	save_file.write(upload.file.read())
	save_file.close()
	return name, ext, name_hash+ext

@app.route('/to/pdf', method='POST')
def convert_to_pdf():
	orig_word_filename, ext, word_filename = save_word_file()
	word_file_path = UPLOAD_PATH+'\\'+word_filename
	item = WorkQueueItem(word_file_path, word_filename, orig_word_filename, ext, 0)
	item.wait_work() # Espera por la conversion
	os.remove(word_file_path)
	return static_file(word_filename+'.pdf', root=PDF_PATH, download=orig_word_filename+'.pdf')

@app.route('/cleanup/word', method='POST')
def cleanup_word():
	orig_word_filename, ext, word_filename = save_word_file()
	word_file_path = UPLOAD_PATH+'\\'+word_filename
	item = WorkQueueItem(word_file_path, word_filename, orig_word_filename, ext, 1)
	item.wait_work() # Espera por la limpieza
	GCFile(word_file_path)
	return static_file(word_filename, root=UPLOAD_PATH, download=orig_word_filename+ext)

def main():
	worker_thread = WorkQueueThread()
	gc=threading.Thread(target=gc_thread)
	gc.start()
	worker_thread.start()
	run(app, server='paste', host='0.0.0.0', port=8080)

if __name__ == "__main__":
	main()
