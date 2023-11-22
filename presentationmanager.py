import os
import copy
import tempfile
import traceback

from pathlib import Path
from pptx import Presentation
from pptx.shapes.graphfrm import GraphicFrame
from pptx.shapes.picture import Picture
from pptx.shapes.group import GroupShape

from common import create_text_chunks

class PresentationManager(object):
    """Contains Presentation object and functions to manage it"""
    
    # Character limit for content text in single slide
    MAX_CONTENT_LIMIT=2250

    def __init__(self, file_path=None, template_slide_index=1, slide_size=()):
        # Since presentation.Presentation class not intended to be constructed directly, using pptx.Presentation() to open presentation
        if file_path and Path(file_path).exists():
            self.presentation = Presentation(file_path)
            print("Loaded presentation from:", file_path)
        else:
            if file_path:
                print(f"Could not load {file_path}")  
            self.presentation = Presentation()
            print("New presentation object loaded")

        if slide_size:
            height, width = slide_size
            self.presentation.slide_height = height
            self.presentation.slide_width = width

        # Setting index of slide to be used as a template
        self.template_slide_index = template_slide_index

        # Get index of Blank slide layout
        layout_items_count = [len(layout.placeholders) for layout in self.presentation.slide_layouts]
        min_items = min(layout_items_count)
        self.blank_layout_id = layout_items_count.index(min_items)

    @property
    def xml_slides(self):
        return self.presentation.slides._sldIdLst

    @property
    def _blank_slide_layout(self):        
        return self.presentation.slide_layouts[self.blank_layout_id]
    
    @property
    def total_slides(self):
        return len(self.presentation.slides)

    def duplicate_slide(self, index, destination=None):
        """
        Duplicates the slide with the given index. Adds slide to the end of the presentation
        """
        source = self.presentation.slides[index]
        destination = destination or self
        # Adds blank slide to end
        blank_slide_layout = destination._blank_slide_layout
        dest = destination.presentation.slides.add_slide(blank_slide_layout)

        # Creates empty list and empty folder `temp` in project
        images = {}

        with tempfile.TemporaryDirectory() as tmpdirname:
            for shp in source.shapes:

                if isinstance(shp, GroupShape):
                    for grp_shape in shp.shapes:
                        if isinstance(grp_shape, Picture):
                            # Save image
                            filepath = os.path.join(tmpdirname, grp_shape.name+'.jpg')
                            with open(filepath, 'wb') as f:
                                f.write(grp_shape.image.blob)
                            # Add image path and size to dict `images`
                            images[filepath] = [grp_shape.left, grp_shape.top, grp_shape.width, grp_shape.height]                        

                if isinstance(shp, Picture):
                    # Save image
                    filepath = os.path.join(tmpdirname, shp.name+'.jpg')
                    with open(filepath, 'wb') as f:
                        f.write(shp.image.blob)
                    # Add image path and size to dict `images`
                    images[filepath] = [shp.left, shp.top, shp.width, shp.height]

                

                
                # Add all other slide elements
                if not isinstance(shp, GraphicFrame):
                    el = shp.element
                    newel = copy.deepcopy(el)
                    dest.shapes._spTree.insert_element_before(newel, 'p:extLst')
        
            # Add images to new slide and remove from filesystem
            for k, v in images.items():
                dest.shapes.add_picture(k, v[0], v[1], v[2], v[3])

        return dest

    def move_slide(self, old_index, new_index):
        slides = list(self.xml_slides)
        self.xml_slides.remove(slides[old_index])
        self.xml_slides.insert(new_index, slides[old_index])

    def remove_slide(self, index):
        slides = list(self.xml_slides)
        self.xml_slides.remove(slides[index]) 

    def remove_all_slides(self):
        slides = list(self.xml_slides)
        for slide in slides:
            self.xml_slides.remove(slide)       
        

    def add_text_to_slide(self, index, text_content, title=""):
        """Adds title and content to slide at given index"""
        
        dest = self.presentation.slides[index]
        # Get title frame and content frame
        text_frames = []
        for shape in dest.shapes:
            if shape.has_text_frame:
                if 'Title' in shape.name:
                    title_frame = shape.text_frame
                else:
                    text_frames.append(shape.text_frame)
        # Choose first text frame as target
        content_frame = text_frames[0]
        
        # Clear content frame and add text
        content_frame.clear()
        p = content_frame.paragraphs[0]
        run = p.add_run()
        run.text = text_content

        # Clear title frame and add title
        title_frame.clear()
        p = title_frame.paragraphs[0]
        run = p.add_run()
        run.text = title

    def populate_slide(self, content, title=""):
        """Creates slides with given text and title, making more slides if text over limit"""
        
        duplicate_indices = []
        chunks = create_text_chunks(content, self.MAX_CONTENT_LIMIT)

        # Create slides for each chunk of text
        for chunk in chunks:
            slide_copy = self.duplicate_slide(self.template_slide_index)
            i = self.presentation.slides.index(slide_copy)
            duplicate_indices.append(i)
            self.add_text_to_slide(i, chunk, title)

        # Move all slides to just before template slide
        new_index = self.template_slide_index + 1
        for old_index in duplicate_indices:
            self.move_slide(old_index, new_index)
            new_index += 1


    def save(self, filepath, remove_template=False):
        """Saves presentation to given filepath and removes slide used as template"""

        if remove_template:
            print("Removing template", self.template_slide_index)
            self.remove_slide(self.template_slide_index)
        self.presentation.save(filepath)
        print("Saved presentation to:", filepath)


    @classmethod
    def copy_slide_to_other_presentation(cls, source, dest_filepath, slides_to_copy=[]):
        height, width = source.presentation.slide_height, source.presentation.slide_width
        destination = PresentationManager(dest_filepath, slide_size=(height, width))
        if not slides_to_copy:
            slides_to_copy = range(source.total_slides)
        try:
            for i in slides_to_copy:
                source.duplicate_slide(i, destination)
            destination.save(dest_filepath)
            destination = Presentation(dest_filepath)
            destination.save(dest_filepath)
        except Exception:
            traceback.print_exc()   
