'''
David DeFilippis 2021-05-25
GUI program to get coordinates from a stem map and save them as an excel file.
The program is designed to be lightweight and use as few dependencies as possible.
'''

# third party dependencies
import PySimpleGUI as sg
from PIL import Image, ImageGrab
import xlsxwriter
# dependencies that are part of the standard library
import re
import os.path
import io
import base64



"""
@TODO
Save the images with the user input overlays
Check that the different window sizes work properly
Package up the progect and get it on github

Small screen sizes cut off the erase last point button. 
Not a big enough issue for now.

@NOTE

"""

# start the program by getting the size of the screen
# we will use this to select the size fo the font which, in turn sets the size of the elements
screen_size = sg.Window.get_screen_size()
print(screen_size)
# medium screen resolution
# screen_size = (1366, 786)
# small screen resolution
# screen_size = (1024, 600)
# we are mostly concerned with the 
if screen_size[1] >= 1080:
    fontsize= 14
    radio_fontsize = fontsize
    table_rows = 20
    point_size = 12
    label_input_size = 15
    # set the canvas size based on the screen size, leaving room for the other column and a little room for the radio buttons below
    canvas_size = (screen_size[0] - 300, screen_size[1]-120)
    # since the image is 8.5 x 11 we can have the image zoomed in a little bit
    image_size = (1400, 1400)
elif screen_size[1] < 1080 and screen_size[1] >= 720:
    fontsize = 11
    radio_fontsize = fontsize - 1
    table_rows = 10
    point_size = 8
    label_input_size = 10
    # set the canvas size based on the screen size, leaving room for the other column and a little room for the radio buttons below
    canvas_size = (screen_size[0] - 300, screen_size[1]-100)
    image_size = (950, 950)
else:
    fontsize = 11
    radio_fontsize = fontsize - 2
    table_rows = 5
    point_size = 6
    label_input_size = 5
    # set the canvas size based on the screen size, leaving room for the other column and a little room for the radio buttons below
    canvas_size = (screen_size[0] - 300, screen_size[1]-100)
    image_size = (700, 700)

# holds the language selection value english is 1 and spanish is 0
is_english = 1
font = ('menlo', fontsize)
# location is based off of the top left corner, instead of the bottom left, for some dumb reason
image_loc = (0, canvas_size[1]*1.25)
# spinbox needs to have a list of possible values so we can make quadrat sizes anywhere from 0 - 100 meters in length
quadrat_size_values = [ i for i in range(100)]
# make the table header
table_header = [' type ', ' label ', '  x  ', '  y  ', ' local x ', ' local y ', ' plot x ', ' plot y ']
# variable to hold data for mouse click coordinates in the table
coordinate_list = [
    ['NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA'],
    ['NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA'],
    ['NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA'],
    ['NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA'],
    ['NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA'],
    ['NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA'],
    ['NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA'],
    ['NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA'],
    ['NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA'],
    ['NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA']
]

# pixel points to meters conversion
pixels_per_meter = 0
origin_pixels = (0, 0)
# standard quadrat sides are 20 meters
initial_quadrat_size = 20
point_counter = 3
# keep track of all the points so we can delete them if we need to
points = {}
# label to update point @NOTE this variable is not always an integer with lianas due to clonal stem designation
point_label = 0
# dictionaries to hold the text values for elements in english and spanish
text_element_lang_dict = {
    '-FOLDER TXT-' : ('Carpeta', ' Folder'),
    '-QUAD NUM TXT-' : ('Numero Cuadrante', '  Quadrat number   '),
    '-QUAD SIZE TXT-' : ('Tamaño del cuadrante (m)', '    Quadrat size (m)     '),
    '-RESIZE TEXT-' : ('Tamaño de la imagen', 'Resize the image to'),
    '-SAVE FOLDER TXT-' : ('Guardar carpeta', '  Save Folder  '),
    '-SAVE NAME TXT-' : ('Nombre archivo', ' File name '),
    '-LABEL TXT-' : ('Etiqueta', '  Label ')
}
button_element_lang_dict = {
    '-RESIZE BTN-' : ('Cambiar el tamaño', 'Resize'),
    '-BRZ BTN-' : ('Buscar', 'Browse'),
    '-SAVE BRWS-' : ('Buscar', 'Browse' ),
    '-SAVE BTN-' : ('Guardar', 'Save'),
    '-POINT UPDT BTN-' : ('Actualizar', 'Update'),
    '-CLEAR BTN-' : ('Borrar punto', 'Erase Point')
}
radio_element_lang_dict = {
    '-IMG MOVE-' : ('Mover Imagen', ' Move Image '),
    '-RECT-' : ('Dibujar contorno', '  Draw Outline  '),
    '-POINT-' : ('Elegir puntos', '  Select Points   ')
    }

def switch_language(val : int):
    if val:
        for k, v in text_element_lang_dict.items():
            window[k].update(value=v[1])
        for k, v in button_element_lang_dict.items():
            window[k].update(v[1])
        for k, v in radio_element_lang_dict.items():
            window[k].update(text=v[1])
    else:
        for k, v in text_element_lang_dict.items():
            window[k].update(value=v[0])
        for k, v in button_element_lang_dict.items():
            window[k].update(v[0])
        for k, v in radio_element_lang_dict.items():
            window[k].update(text=v[0])


# Functions - we may break these out into a seperate module later.
def parse_quad_number(input : str) -> str:
    """ get the integer from the filename """
    var = re.search(r'\d+', input)
    if var.group():
        return var.group()
    else:
        return '0000'

def make_directory(input_path : str):
	"""make sure that the directory we want to save to exists, if not make it"""
	import os # moving around the file system
	if not os.path.exists(input_path): # if the path doesn't exist
		os.makedirs(input_path) # lets make it
		window["-INFO-"].update(value=f"Directory {input_path} Created ")
	else:
		window["-INFO-"].update(value=f"Directory {input_path} already exists")

def table_data_to_excel(data : list, path : str, filename : str, analyst : str):
    """ Open an excel workbook and write the table data to it """
    # check that the path exists, if not create it
    make_directory(path)
    table_header.append('Entry Analyst')
    window["-INFO-"].update(value=values['-POINT TABLE-'])
    # write the table data to the file
    with xlsxwriter.Workbook(f'{path}/{filename}.xlsx') as workbook:
        worksheet = workbook.add_worksheet()
        # write in the header
        worksheet.write_row(0, 0, table_header)
        # wirte in all the rows of data
        for idx, row in enumerate(data):
            row.append(analyst) # add the data entry analyst to the last column in the row
            worksheet.write_row(idx+1, 0, row) # write the row to the excel file
    # workbook.close() # warn("Calling close() on already closed file.") they say
    table_header.pop() # remove the data entry analyst column

def save_element_as_file(element, filename):
    """
    Saves any element as an image file.  Element needs to have an underlyiong Widget available (almost if not all of them do)
    :param element: The element to save
    :param filename: The filename to save to. The extension of the filename determines the format (jpg, png, gif, ?)
    """
    widget = element.Widget
    box = (widget.winfo_rootx(), widget.winfo_rooty(), widget.winfo_rootx() + widget.winfo_width(), widget.winfo_rooty() + widget.winfo_height())
    grab = ImageGrab.grab(bbox=box)
    grab.save(filename)

def get_pixels_to_meters(rect_coords : tuple, side_len : int) -> float:
    """ Uses the rectangle coordinates ((min_x, min_y), (max_x, max_y)) and the length of the side in meters to get the pixels per meter. """
    pix_per_met = []
    try:
        pix_per_met.append((rect_coords[1][0] - rect_coords[0][0]) / side_len)
        pix_per_met.append((rect_coords[1][1] - rect_coords[0][1]) / side_len)
        return sum(pix_per_met)/len(pix_per_met)
    except ZeroDivisionError:
        return 0

def get_local_coordinates(pixel_coords : tuple, origin_coords : tuple, pix_per_met : float) -> tuple:
    """ get teh local coordinates from the pixel coordinates of a point """
    delta_pixels = (pixel_coords[0] - origin_coords[0], pixel_coords[1] - origin_coords[1])
    try:
        return (round(delta_pixels[0] / pix_per_met, 2), round(delta_pixels[1] / pix_per_met, 2))
    except ZeroDivisionError:
        return (0,0)

def get_plot_coordinates(local_coords : tuple, quad : int, side_len_meters : int) -> tuple:
    """ Return the plot level coordinates based on the quadrat number and the size of the quadrat at the plot level.
    :param local_coords: the local coordinates for the point
    :param quad: the quadrat number in the plot
    :param side_len_meters the length of the sides of the quadrat
      """
    # make sure the quad number is an integer
    quad = int(quad)
    # get the horizontal position of the quadrat
    horz = int(quad/100)
    # get the vertical position of the quadrat
    vert = int(quad%100)
    # calculate and return the plot level coordinates of the point
    return (horz * side_len_meters + local_coords[0], vert * side_len_meters + local_coords[1])

def convert_to_bytes(file_or_bytes, resize=None):
    '''
    Will convert into bytes and optionally resize an image that is a file or a base64 bytes object.
    Turns into  PNG format in the process so that can be displayed by tkinter
    :param file_or_bytes: either a string filename or a bytes base64 image object
    :type file_or_bytes:  (Union[str, bytes])
    :param resize:  optional new size
    :type resize: (Tuple[int, int] or None)
    :return: (bytes) a byte-string object
    :rtype: (bytes)
    '''
    if isinstance(file_or_bytes, str):
        img = Image.open(file_or_bytes)
    else:
        try:
            img = Image.open(io.BytesIO(base64.b64decode(file_or_bytes)))
        except Exception as e:
            dataBytesIO = io.BytesIO(file_or_bytes)
            img = Image.open(dataBytesIO)

    cur_width, cur_height = img.size
    if resize:
        new_width, new_height = resize
        try:
            scale = min(new_height/cur_height, new_width/cur_width)
        except ZeroDivisionError:
            scale = 0
        img = img.resize((int(cur_width*scale), int(cur_height*scale)), Image.ANTIALIAS)
    bio = io.BytesIO()
    img.save(bio, format="PNG")
    del img
    return bio.getvalue()



# --------------------------------- Define Layout Columns ---------------------------------
sg.theme('Dark Grey 1')


# First the window layout...2 columns

left_col = [
    # output infor for the developer
    [sg.Text(key='-INFO-', size=(40, 4))],
    # [sg.Text('Input file path:', key='-INFO TEXT-'), sg.Text(key='-TOUT-', size=(20, 1))],
    [
        sg.Text('Español', key='-ESP LANG-'),
        sg.Slider(
            range=(0,1), 
            default_value=1, 
            resolution=1, 
            size=(14, None), 
            disable_number_display = True,
            orientation='horizontal', 
            background_color = 'green',
            font=font, 
            key='-LANG SLIDER-',
            enable_events=True
        ),
        sg.Text('English', key='-ENG LANG-'),
        sg.In('Data Analyst', size=(30, None), enable_events=True, key='-DEA-')
    ],
    [sg.HorizontalSeparator()],
    # file selection
    [sg.Text(text_element_lang_dict['-FOLDER TXT-'][1], key='-FOLDER TXT-'), sg.In(enable_events=True ,key='-FOLDER-'), sg.FolderBrowse(auto_size_button=True, key='-BRZ BTN-')],
    [sg.Listbox(values=[], enable_events=True, size=(40,table_rows),key='-FILE LIST-')],
    [sg.HorizontalSeparator()],
    [
        sg.Text(text_element_lang_dict['-QUAD SIZE TXT-'][1], key='-QUAD SIZE TXT-'), 
        sg.Spin(values=quadrat_size_values, initial_value=initial_quadrat_size, size=(5,None), auto_size_text=True, key='-QUAD SIZE-'), 
        sg.Text(text_element_lang_dict['-QUAD NUM TXT-'][1], key='-QUAD NUM TXT-'), sg.In('0000', key='-QUAD NUM-', size=(5, 1))
    ],
    [sg.HorizontalSeparator()],
    [
        sg.Text(text_element_lang_dict['-RESIZE TEXT-'][1], key='-RESIZE TEXT-'), 
        sg.In(key='-W-', size=(5,1)), 
        sg.In(key='-H-', size=(5,1)), 
        sg.Button('Resize', size=(None,1), key='-RESIZE BTN-')
    ],        
    [sg.Table(
        values=coordinate_list,
        headings=table_header,
        def_col_width = 10,
        auto_size_columns=True,
        max_col_width = 50,
        num_rows = table_rows,
        display_row_numbers=True,
        justification='left',
        alternating_row_color='green',
        key='-POINT TABLE-',
        enable_events=True,
        font=('menlo', fontsize-1)
        )
    ],
    [sg.HorizontalSeparator()],
    [
        sg.Text(text_element_lang_dict['-SAVE FOLDER TXT-'][1], key='-SAVE FOLDER TXT-'), 
        sg.In(size=(40,None), key='-SAVE LOC-', justification='left', enable_events=True), 
        sg.FolderBrowse(key='-SAVE BRWS-')], 
    [
        sg.Text(text_element_lang_dict['-SAVE NAME TXT-'][1], key='-SAVE NAME TXT-'), 
        sg.In(key='-SAVE NAME-', enable_events=True), 
        sg.Button('Save', size=(None, 1), key='-SAVE BTN-', enable_events=True)]
]


image_col = [
    [
        sg.Graph(
            # make graph bigger than canvas so we can move things off the top and left of canvas
            canvas_size=(canvas_size[0], canvas_size[1]),
            graph_bottom_left=(0, 0),
            # this essentially controls the graph resolution
            graph_top_right=(canvas_size[0], canvas_size[1]),
            key="-GRAPH-",
            background_color = 'gray',
            enable_events=True,
            drag_submits=True
        )
    ],
    [sg.HorizontalSeparator()],
    [
        sg.Text(text_element_lang_dict['-LABEL TXT-'][1], key='-LABEL TXT-'), 
        sg.In(size=(label_input_size, None), enable_events=True ,key='-LABEL-'), 
        sg.Button('Update', key='-POINT UPDT BTN-'),
        sg.VerticalSeparator(),
        sg.R('Move Image', 1, font=(font[0], radio_fontsize), key='-IMG MOVE-', enable_events=True, default=True),
        sg.R('Draw Outline', 1, font=(font[0], radio_fontsize), key='-RECT-', enable_events=True), 
        sg.R('Select Points', 1,  font=(font[0], radio_fontsize), key='-POINT-', enable_events=True), 
        sg.Button('Erase Point', key='-CLEAR BTN-', enable_events=True)
    ],
]

# ----- Full layout -----
layout = [[sg.Column(left_col, element_justification='c'), sg.VSeperator(),sg.Column(image_col, element_justification='c')]]

# --------------------------------- Create Window ---------------------------------

window = sg.Window("Stem Coordinate Entry Assistant", layout, font=font, resizable=False, size=screen_size)

graph = window.Element('-GRAPH-')

window.Finalize().FindElement('-GRAPH-').Widget.config(cursor="X_cursor")
# make a couple of globals to hold some status values
dragging = False
start_point = end_point = prior_rect = None

quadrat_size = initial_quadrat_size

# holds onto the id of the imported image
image = 0

# print(window.AllKeysDict)

# ----- Run the Event Loop -----
# --------------------------------- Event Loop ---------------------------------
while True:
    event, values = window.read()
    
    if event in (sg.WIN_CLOSED, 'Exit'):
        break
    if event == sg.WIN_CLOSED or event == 'Exit':
        break

    # ----- Get the data here -----
    # if  we hit the folder button and have selected a folder path
    if event == '-FOLDER-':                         # Folder name was filled in, make a list of files in the folder
        # get the path to the folder
        folder = values['-FOLDER-']
        # try and parse the files within the folder
        try:
            file_list = os.listdir(folder)         # get list of files in folder
        except:
            file_list = []
        # get the list of files that are one of the image formats
        fnames = [f for f in file_list if os.path.isfile(
            os.path.join(folder, f)) and f.lower().endswith((".png", ".jpg", "jpeg", ".tiff", ".bmp"))]
        # update the file list in the multitext box
        window['-FILE LIST-'].update(fnames)
        window['-FOLDER-'].Widget.xview("end")
    # else if a file was chosen from the listbox
    elif event == '-FILE LIST-':
        # try and get the filename
        try:
            filename = os.path.join(values['-FOLDER-'], values['-FILE LIST-'][0])
            # window['-TOUT-'].update(filename)
            if values['-W-'] and values['-H-']:
                new_size = int(values['-W-']), int(values['-H-'])
            else:
                new_size = (image_size[0], image_size[1])
            #window['-IMAGE-'].update(data=convert_to_bytes(filename, resize=new_size))
            # draw the image on the graph
            image = graph.DrawImage(data=convert_to_bytes(filename, resize=new_size), location=image_loc)

            # we can use the input name to pre-fill some of the other 
            quad_number = parse_quad_number(values['-FILE LIST-'][0])
            window['-QUAD NUM-'].update(quad_number)
            window['-SAVE NAME-'].update(f'Q{quad_number}')
            window['-SAVE LOC-'].update(folder)
            window['-SAVE LOC-'].Widget.xview("end")
        except Exception as E:
            sg.popup_error(f'** Error getting file from list {E} **')
            pass        # something weird happened making the full filename

    # ----- Save the data here -----
    # if we clicked the save button
    # check if there is a file and a path
    if event == '-SAVE BTN-':
        if values['-SAVE NAME-'] and values['-SAVE LOC-']:

            window["-INFO-"].update(value=f"Saving table to {values['-SAVE LOC-']} as {values['-SAVE NAME-']}.xlsx")
            table_data_to_excel(coordinate_list, values['-SAVE LOC-'], values['-SAVE NAME-'], values['-DEA-'])
            # save an image of the quadrat with the drawings on top
            # save_element_as_file(window['-GRAPH-'], f"{values['-SAVE LOC-']}/{values['-SAVE NAME-']}_img.p")
            window["-INFO-"].update(value=f"Saved table to {values['-SAVE LOC-']} as {values['-SAVE NAME-']}.xlsx")
            # need to do something here that clears everything up so there is a sense of starting over again.
            # clear and update table
            window['-POINT TABLE-'].update(values=[['NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA']])
            coordinate_list = [['NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA']]
            # clear the label value input box
            window['-LABEL-'].update(value='')
            # select the image move radio button
            window['-IMG MOVE-'].update(value=True)
            # clear graph
            graph.erase()
            # select next file in file list
        else:
            window["-INFO-"].update(value=f"No path or filename given")

    # -----If the quadrat size spinner changes values update it here -----
    if event == '-QUAD SIZE-':
        quadrat_size = values['-QUAD SIZE-']
    if event == '-LANG SLIDER-':
        switch_language(values['-LANG SLIDER-'])
    # if we pressed the button to resize the image
    if event == '-RESIZE BTN-': 
        # get the new values from the input boxex
        if values['-W-'] and values['-H-']:
            new_size = int(values['-W-']), int(values['-H-'])
        else:
            new_size = None
        # try and get the origin of the drawn rectangle
        # print(origin_pixels)
        if origin_pixels:
            # delete the old image
            # for figure in drag_figures:
            #         graph.delete_figure(figure)
            # load in the new image with the new size
            print(f"image id we are trying to delete: {image}")
            graph.delete_figure(image)
            image = graph.DrawImage(data=convert_to_bytes(filename, resize=new_size), location=image_loc)
        else:
            # for figure in drag_figures:
            #         graph.delete_figure(figure)
            image = graph.DrawImage(data=convert_to_bytes(filename, resize=new_size), location=image_loc)
    # if we hit the erase button erase everything @NOTE still need to get this working
    if event == '-CLEAR BTN-':
        # delete the last point drawn on the figure
        graph.delete_figure(points[max(points.keys())])
        # pop the last row from the coordinate list
        coordinate_list.pop()
        # update the point table
        window['-POINT TABLE-'].update(values=coordinate_list)
        # debug
        window['-INFO-'].update(value=points)
    # make button to erase last point from graph and point table

    # ------ Lets try and draw on the image ------
    # if there's a "Graph" event, then it's a mouse click on the graph
    if event == "-GRAPH-":
        # get the coodinates of the mouse cursor
        x, y = values["-GRAPH-"]
        # if we are not dragging a mouse click then update the mouse cursor location as the start point
        if not dragging:
            start_point = (x, y)
            dragging = True
            drag_figures = graph.get_figures_at_location((x,y))
            lastxy = x, y
        # else if we are dragging then update the end point as the mouse cursor location
        else:
            end_point = (x, y)
        # if we are drawing a new rectangle delete the previous one
        if prior_rect:
            graph.delete_figure(prior_rect)
        # update the coordinates
        delta_x, delta_y = x - lastxy[0], y - lastxy[1]
        lastxy = x,y
        # if we have values in the start and end points - double negative here
        if None not in (start_point, end_point):
            if values['-IMG MOVE-']:
                for fig in drag_figures:
                    graph.move_figure(fig, delta_x, delta_y)
                    graph.update()
            # if we are drawing the bounding rectangle
            elif values['-RECT-']:
                # draw the rectangle while we have the left mouse button held down
                prior_rect = graph.draw_rectangle(start_point, end_point,fill_color=None, line_color='green', line_width=3)
                # create the first two rows of the points table
                # convert the pixel coordinates on the map to meters
                origin_pixels = start_point
                pixels_per_meter = get_pixels_to_meters((start_point, end_point), quadrat_size)

                start_meters = get_local_coordinates(start_point, origin_pixels, pixels_per_meter)
                end_meters = get_local_coordinates(end_point, origin_pixels, pixels_per_meter)

                start_global = get_plot_coordinates(start_meters, values['-QUAD NUM-'], quadrat_size)
                end_global = get_plot_coordinates(end_meters, values['-QUAD NUM-'], quadrat_size)

                coordinate_list = [
                    ['origin', 'p0', start_point[0], start_point[1], start_meters[0], start_meters[1], start_global[0], start_global[1]], 
                    ['maximum', 'p1', end_point[0], end_point[1], end_meters[0], end_meters[1], end_global[0], end_global[1]]
                ]
                # save our origin for resizing the image
                # origin_pixels = (start_point[0], end_point[0])
            elif values['-POINT-']:
                points[point_counter] = graph.draw_point((x,y), color='green', size=point_size)
                point_meters = get_local_coordinates((x, y), origin_pixels, pixels_per_meter)
                point_global = get_plot_coordinates(point_meters, values['-QUAD NUM-'], quadrat_size)
                # point_global = (0, 0)
                coordinate_list.append(['point', point_counter, x, y, point_meters[0], point_meters[1], point_global[0], point_global[1]])
                point_counter = point_counter + 1
                # window['-POINT TABLE-'].update(values=coordinate_list)
        window["-INFO-"].update(value=f"dragging mouse at coordinates {values['-GRAPH-']}")

    # when mouse click is released, update the point table 
    elif event.endswith('+UP'):  # The drawing has ended because mouse up
        window["-INFO-"].update(value=f"grabbed rectangle from {start_point} to {end_point}")
        window['-POINT TABLE-'].update(values=coordinate_list)
        # scroll the table to the bottom to always see the latest entry
        window['-POINT TABLE-'].Widget.yview_moveto(1)
        start_point, end_point = None, None  # enable grabbing a new rect
        dragging = False
        prior_rect = None

    # # conditions to mouse actions: right clicking on mouse and other mouse cursor motion
    # elif event.endswith('+RIGHT+'):  # Righ click
    #     window["-INFO-"].update(value=f"Right clicked location {values['-GRAPH-']}")
    # elif event.endswith('+MOTION+'):  # Righ click
    #     window["-INFO-"].update(value=f"mouse freely moving {values['-GRAPH-']}")

    # --------- update the last coordinates label -------------
    if event == '-POINT UPDT BTN-':
        # get the last row from the coordinate list and update the second element with the label input box
        coordinate_list[-1][1] = values['-LABEL-']
        # push the update to the point table
        window['-POINT TABLE-'].update(values=coordinate_list)

# --------------------------------- Close & Exit ---------------------------------
window.close()