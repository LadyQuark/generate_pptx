from pptx import Presentation
import os
import copy
from pathlib import Path
from common import create_text_chunks

class PresentationManager(object):
    """Contains Presentation object and functions to manage it"""
    
    # Character limit for content text in single slide
    MAX_CONTENT_LIMIT=2250

    def __init__(self, file_path, template_slide_index=1):
        # Since presentation.Presentation class not intended to be constructed directly, using pptx.Presentation() to open presentation
        self.presentation = Presentation(file_path)
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
    def _get_blank_slide_layout(self):        
        return self.presentation.slide_layouts[self.blank_layout_id]

    def duplicate_slide(self, index):
        """
        Duplicates the slide with the given index. Adds slide to the end of the presentation
        """
        source = self.presentation.slides[index]
        # Adds blank slide to end
        blank_slide_layout = self._get_blank_slide_layout()
        dest = self.presentation.slides.add_slide(blank_slide_layout)

        # Creates empty list and empty folder `temp` in project
        images = {}
        temp_folder = "temp"
        Path(temp_folder).mkdir(parents=True, exist_ok=True)
        # all images in slide
        for shp in source.shapes:
            if 'Picture' in shp.name:
                # Save image to folder `temp`
                filepath = os.path.join(temp_folder, shp.name+'.jpg')
                with open(filepath, 'wb') as f:
                    f.write(shp.image.blob)
                # Add image path and size to dict `images`
                images[filepath] = [shp.left, shp.top, shp.width, shp.height]
        
        # Add images to new slide and remove from filesystem
        for k, v in images.items():
            dest.shapes.add_picture(k, v[0], v[1], v[2], v[3])
            os.remove(k)

        # Add all other slide elements
        for shp in source.shapes:
            if 'Picture' not in shp.name:
                el = shp.element
                newel = copy.deepcopy(el)
                dest.shapes._spTree.insert_element_before(newel, 'p:extLst')

        return dest

    def move_slide(self, old_index, new_index):
        slides = list(self.xml_slides)
        self.xml_slides.remove(slides[old_index])
        self.xml_slides.insert(new_index, slides[old_index])

    def remove_slide(self, index):
        slides = list(self.xml_slides)
        self.xml_slides.remove(slides[index]) 
        

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

    def save(self, filepath, remove_template=True):
        """Saves presentation to given filepath and removes slide used as template"""

        if remove_template:
            print("Removing template", self.template_slide_index)
            self.remove_slide(self.template_slide_index)
        self.presentation.save(filepath)


