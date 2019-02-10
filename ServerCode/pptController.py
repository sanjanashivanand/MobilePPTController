import win32com.client
import win32api
import win32con
import pythoncom

VK_CODE = {
	'spacebar':0x1B,
	'down_arrow':0x28,
}
class pptController:
	def __init__(self):    # looks for the powerpoint thats already open
		pythoncom.CoInitialize()
		self.app = win32com.client.Dispatch("PowerPoint.Application")

	def fullScreen(self):   # to make fullscreen mode

		if self.hasActivePresentation():
			self.app.ActivePresentation.SlideShowSettings.Run()
			return self.getActivePresentationSlideIndex()

	def click(self):	# to exit from the full screen mode
		win32api.keybd_event(VK_CODE['spacebar'],0,0,0)
		win32api.keybd_event(VK_CODE['spacebar'],0,win32con.KEYEVENTF_KEYUP,0)
		return self.getActivePresentationSlideIndex()

	def gotoSlide(self,index):	# to go to the specific slide

		if self.hasActivePresentation():
			try:
				self.app.ActiveWindow.View.GotoSlide(index)
				return self.app.ActiveWindow.View.Slide.SlideIndex
			except:
				self.app.SlideShowWindows(1).View.GotoSlide(index)
				return self.app.SlideShowWindows(1).View.CurrentShowPosition

	def nextPage(self):		# go to next page
		if self.hasActivePresentation():
			count = self.getActivePresentationSlideCount()
			index = self.getActivePresentationSlideIndex()
			return index if index >= count else self.gotoSlide(index+1)

	def prePage(self):		# go to previous page
		if self.hasActivePresentation():
			index =  self.getActivePresentationSlideIndex()
			return index if index <= 1 else self.gotoSlide(index-1)

	def getActivePresentationSlideIndex(self):  # display current slide number

		if self.hasActivePresentation():
			try:
				index = self.app.ActiveWindow.View.Slide.SlideIndex
			except:
				index = self.app.SlideShowWindows(1).View.CurrentShowPosition
		return index

	def getActivePresentationSlideCount(self):		# display total file count
		print self.app.ActivePresentation.Slides.count
		return self.app.ActivePresentation.Slides.Count

	def getPresentationCount(self):		# count the total slide
		return self.app.Presentations.Count

	def hasActivePresentation(self):		# check whether there is any more slides
		return True if self.getPresentationCount() > 0 else False


