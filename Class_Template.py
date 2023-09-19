@property  # Get
def pdf_iframe_element(self):
 return self._pdf_iframe_element
 # return self._name

@pdf_iframe_element.setter  # Set
def pdf_iframe_element(self, pdf_iframe_element):
 """Set the property."""
 self._pdf_iframe_element = pdf_iframe_element

@pdf_iframe.deleter  # Delete
def name(self):
 """Delete the dog's name."""
 del self._pdf_iframe

@name.setter
def name(self, new_name):
 """Set the dog's name."""
 self._name = new_name

@name.deleter
def name(self):
 """Delete the dog's name."""
 del self._name