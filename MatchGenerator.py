import ast
import base64
import ctypes
import os
import random
import sys
import winreg
from pathlib import Path

import PySimpleGUI as sg

# ----- SETUP ----- #

rigth_click_menu = ['&Right', ['Clear', 'Select hero', 'Edit properties']]
cell_size = 4
margin = 3
columns = 6
rows = 3
positions = ["Vector(10, 0, -10)", "Vector(10, 0, -6)", "Vector(10, 0, -2)", "Vector(10, 0, 2)", "Vector(10, 0, 6)",
             "Vector(10, 0, 10)", "Vector(6, 0, -10)", "Vector(6, 0, -6)", "Vector(6, 0, -2)", "Vector(6, 0, 2)",
             "Vector(6, 0, 6)", "Vector(6, 0, 10)", "Vector(2, 0, -10)", "Vector(2, 0, -6)", "Vector(2, 0, -2)",
             "Vector(2, 0, 2)", "Vector(2, 0, 6)", "Vector(2, 0, 10)", "Vector(-10, 0, 10)", "Vector(-10, 0, 6)",
             "Vector(-10, 0, 2)", "Vector(-10, 0, -2)", "Vector(-10, 0, -6)", "Vector(-10, 0, -10)",
             "Vector(-6, 0, 10)", "Vector(-6, 0, 6)", "Vector(-6, 0, 2)", "Vector(-6, 0, -2)", "Vector(-6, 0, -6)",
             "Vector(-6, 0, -10)", "Vector(-2, 0, 10)", "Vector(-2, 0, 6)", "Vector(-2, 0, 2)", "Vector(-2, 0, -2)",
             "Vector(-2, 0, -6)", "Vector(-2, 0, -10)"]
mana_values = ["0%", "5%", "10%", "15%", "20%", "25%", "30%", "35%", "40%", "45%", "50%", "55%", "60%", "65%", "70%",
               "75%", "80%", "85%", "90%", "95%", "100%"]
level_values = ["1", "2", "3"]
ow_heroes = ['Ana', 'Ashe', 'Baptiste', 'Bastion', 'Brigitte', 'Cassidy', 'D.Va', 'Doomfist', 'Echo', 'Genji', 'Hanzo',
             'Junkrat', 'Lucio', 'Mei', 'Mercy', 'Moira', 'Orisa', 'Pharah', 'Reaper', 'Reinhardt', 'Roadhog', 'Sigma',
             'Soldier76', 'Sombra', 'Symmetra', 'Torbjorn', 'Tracer', 'Widowmaker', 'Winston', 'Wrecking Ball', 'Zarya',
             'Zenyatta']

last_right_clicked_btn = None

# Associates the config icon with the .mgconfig extension. That's why the program needs admin privileges
try:
    key1 = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, "OverwatchMatchGenerator.mgconfig")
    key2 = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, r"OverwatchMatchGenerator.mgconfig\DefaultIcon")
    winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, r"OverwatchMatchGenerator.mgconfig\shell")
    winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, r"OverwatchMatchGenerator.mgconfig\shell\open")
    command = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, r"OverwatchMatchGenerator.mgconfig\shell\open\command")
    winreg.SetValue(key2, "", winreg.REG_SZ, rf"{Path(sys.argv[0]).parent}\Images\Assets\config_icon.png")
    winreg.SetValue(command, "", winreg.REG_SZ, fr'"{sys.argv[0]}" %1')
    shell32 = ctypes.OleDLL('shell32')
    shell32.SHChangeNotify.restype = None
    event = SHCNE_ASSOCCHANGED = 0x08000000
    flags = SHCNF_IDLIST = 0x0000
    shell32.SHChangeNotify(event, flags, None, None)
except:
    pass


def generate_default_metadata():
    """
    :return: Default metadata dict with a different memory address every time
    """

    return {"hero": "", "mana": 0, "level": 1}


def border(elem, clr):
    """
    :param elem: The element to be surrounded with a border
    :param clr: The color of the border
    :return: A column with a background color of the `clr` var with the element `elem` in it
    """

    return sg.Column([[elem]], background_color=clr)


# ----- CONST LAYOUTS ----- #

team_1_grid = [
    [border(sg.Button("", button_color='black', border_width=0, size=(cell_size, cell_size // 2), pad=margin,
                      auto_size_button=False, key=f"-T1_{i}-", right_click_menu=rigth_click_menu,
                      metadata=generate_default_metadata(),
                      image_subsample=3, image_filename=fr"{Path(sys.argv[0]).parent}\Images\Assets\unknown.png"),
            "#156CED") for i in
     range(columns - 1, -1, -1)],
    [border(sg.Button("", button_color='black', border_width=0, size=(cell_size, cell_size // 2), pad=margin,
                      auto_size_button=False, key=f"-T1_{i + columns}-", right_click_menu=rigth_click_menu,
                      metadata=generate_default_metadata(),
                      image_subsample=3, image_filename=fr"{Path(sys.argv[0]).parent}\Images\Assets\unknown.png"),
            "#156CED") for i in
     range(columns - 1, -1, -1)],
    [border(sg.Button("", button_color='black', border_width=0, size=(cell_size, cell_size // 2), pad=margin,
                      auto_size_button=False, key=f"-T1_{i + columns * 2}-", right_click_menu=rigth_click_menu,
                      metadata=generate_default_metadata(), image_subsample=3,
                      image_filename=fr"{Path(sys.argv[0]).parent}\Images\Assets\unknown.png"), "#156CED")
     for i in
     range(columns - 1, -1, -1)]
]

team_2_grid = [
    [border(sg.Button("", border_width=0, button_color='black', size=(cell_size, cell_size // 2), pad=margin,
                      auto_size_button=False, key=f"-T2_{i + columns * 5}-", right_click_menu=rigth_click_menu,
                      metadata=generate_default_metadata(), image_subsample=3,
                      image_filename=fr"{Path(sys.argv[0]).parent}\Images\Assets\unknown.png"), "#e62727")
     for i in
     range(columns)],

    [border(sg.Button("", border_width=0, button_color='black', size=(cell_size, cell_size // 2), pad=margin,
                      auto_size_button=False, key=f"-T2_{i + columns * 4}-", right_click_menu=rigth_click_menu,
                      metadata=generate_default_metadata(), image_subsample=3,
                      image_filename=fr"{Path(sys.argv[0]).parent}\Images\Assets\unknown.png"), "#e62727")
     for i in
     range(columns)],

    [border(sg.Button("", border_width=0, button_color='black', size=(cell_size, cell_size // 2), pad=margin,
                      auto_size_button=False, key=f"-T2_{i + columns * 3}-", right_click_menu=rigth_click_menu,
                      metadata=generate_default_metadata(), image_subsample=3,
                      image_filename=fr"{Path(sys.argv[0]).parent}\Images\Assets\unknown.png"), "#e62727")
     for i in
     range(columns)]

]

main_layout = [
    [sg.Frame("", team_1_grid, border_width=0)],
    [sg.Button("", key="-RANDOMIZE_DECK-",
               image_filename=rf"{Path(sys.argv[0]).parent}\Images\Assets\randomize_deck.png", image_subsample=3,
               border_width=0, button_color=sg.theme_background_color()),
     sg.Button("", key="-EXPORT-", image_filename=rf"{Path(sys.argv[0]).parent}\Images\Assets\export.png",
               image_subsample=3, border_width=0,
               button_color=sg.theme_background_color())],
    [sg.Button("", key="-CLEAR_DECK-", image_filename=rf"{Path(sys.argv[0]).parent}\Images\Assets\clear_deck.png",
               image_subsample=3,
               border_width=0, button_color=sg.theme_background_color()),
     sg.Button("", key="-IMPORT-", image_filename=rf"{Path(sys.argv[0]).parent}\Images\Assets\import.png",
               image_subsample=3, border_width=0,
               button_color=sg.theme_background_color())],
    [sg.Frame("", team_2_grid, border_width=0)],
]

# The main window which contains all grids and buttons
main_window = sg.Window("ChessWatch Match Generator", main_layout, finalize=True, grab_anywhere=True,
                        icon=rf"{Path(sys.argv[0]).parent}\Images\Assets\logo.ico")


def update_cell_image(elem, hero_name=""):
    """
    Updates a certain cell in a grid with the needed image, tooltip and metadata. This function was created so every
    we need to remove / replace / add a hero to a cell it would be easier

    :param elem: The button in the grid to be updated
    :param hero_name: The hero name for the image, tooltip and metadata of the button
    :return: None
    """

    main_window[elem].metadata["hero"] = hero_name
    main_window[elem].set_tooltip(main_window[elem].metadata["hero"].title())
    if hero_name:
        main_window[elem].update(
            image_filename=f"{Path(sys.argv[0]).parent}\Images/HeroImages\\{main_window[elem].metadata['hero'].lower()}.png",
            image_subsample=3)
    else:
        main_window[elem].update(
            image_filename=fr"{Path(sys.argv[0]).parent}\Images\Assets\unknown.png", image_subsample=3)


def generate_properties_editor_layout(default_settings):
    """
    This will generate a new layout for the properties editor window in a new memory address every time.

    :param default_settings: It will show previous edited properties if there are any
    :return: the layout of the properties editor window
    """

    lvl = 1 if default_settings["level"] == 1 else 2 if default_settings["level"] == 1.25 else 3
    properties_editor_layout = [
        [sg.Button("Change Hero")],
        [sg.Text("Mana:"), sg.Spin(values=mana_values, initial_value=f'{default_settings["mana"]}%', key="-MANA-")],
        [sg.Text("Level:"), sg.Spin(values=level_values, initial_value=lvl, key="-LEVEL-")],
        [sg.Button("Apply")]
    ]
    return properties_editor_layout


def generate_hero_selection_layout():
    """
        This will generate a new layout for the hero selection window in a new memory address every time.

        :return: the layout of the ero selection window
        """

    hero_selection_layout = [
        [sg.Input(do_not_clear=True, size=(20, 1), enable_events=True, key='-SEARCH_INPUT-'),
         sg.Button(image_filename=rf"{Path(sys.argv[0]).parent}\Images\Assets\search.png", image_subsample=12,
                   border_width=0, button_color=sg.theme_background_color(),
                   mouseover_colors=(sg.theme_background_color(), sg.theme_background_color()))],
        [sg.Listbox(values=ow_heroes, select_mode='single', key='-HERO_LIST-', size=(30, 6))],
        [sg.Button("Select")]
    ]
    return hero_selection_layout


def generate_export_settings_layout():
    """
    This will generate a new layout for the export settings window in a new memory address every time.

    :return: the layout of the export settings window
    """

    export_settings_layout = [
        [sg.Button("", key="-EXPORT_TEXT-",
                   image_filename=rf"{Path(sys.argv[0]).parent}\Images\Assets\export_as_txt.png", image_subsample=2,
                   border_width=0, button_color=sg.theme_background_color())],
        [sg.Button("", image_filename=rf"{Path(sys.argv[0]).parent}\Images\Assets\program_code.png", image_subsample=2,
                   border_width=0,
                   button_color=sg.theme_background_color(), key="-PROGRAM_CODE-"),
         sg.Button("", image_filename=rf"{Path(sys.argv[0]).parent}\Images\Assets\overwatch_code.png",
                   image_subsample=2, border_width=0,
                   button_color=sg.theme_background_color(), key="-OVERWATCH_CODE-")]
    ]
    return export_settings_layout


def create_export_settings_window():
    """
    This function contains all the logic of the export settings window and will run the window by it's own.

    :return: None
    """

    window = sg.Window("Export Settings", generate_export_settings_layout(), modal=True, finalize=True)
    while True:
        n_event, n_values = window.read()
        if n_event == "Exit" or n_event == sg.WIN_CLOSED:
            window.close()
            return None

        if n_event == "-PROGRAM_CODE-":
            export(window)

        if n_event == "-OVERWATCH_CODE-":
            export_as_overwatch_code(window)


def create_hero_selection_window():
    """
    This function contains all the logic of the hero selection window and will run the window by it's own.
    
    :return: Selected hero name
    """

    window = sg.Window("Hero selection", generate_hero_selection_layout(), modal=True)
    while True:
        n_event, n_values = window.read()

        if n_event == "Exit" or n_event == sg.WIN_CLOSED:
            break

        # Uses PySimpleGui search in list box example
        if n_values['-SEARCH_INPUT-'] != '':
            search = n_values['-SEARCH_INPUT-']
            new_values = [x for x in ow_heroes if search.lower() in x.lower()]
            window.Element('-HERO_LIST-').Update(new_values)
        else:
            window.Element('-HERO_LIST-').Update(ow_heroes)

        # If they selected a hero it will check if their name should be changed and then returns the hero name
        if n_event == "Select" and len(n_values['-HERO_LIST-']):
            window.close()
            selected_hero = n_values['-HERO_LIST-'][0].lower()
            if "soldier: 76" == n_values['-HERO_LIST-'][0]:
                selected_hero = "soldier76"
            if "McCree" == n_values['-HERO_LIST-'][0]:
                selected_hero = "cassidy"
            return selected_hero

    window.close()


def create_properties_editor_window(selected_hero, btn, settings):
    """
    This function contains all the logic of the properties editor window and will run the window by it's own.

    :param selected_hero: The name of the hero in the selected cell
    :param btn: The key of the button pressed in the grid
    :param settings: Previously edited settings
    :return: updated metadata
    """

    window = sg.Window(f"Hero selection - {selected_hero.title()}", generate_properties_editor_layout(settings),
                       modal=True, size=(200, 150), disable_minimize=True)
    while True:
        n_event, n_values = window.read()
        if n_event == "Exit" or n_event == sg.WIN_CLOSED:
            window.close()
            return None

        if n_event == "Change Hero":
            selected_hero = create_hero_selection_window()
            update_cell_image(btn, selected_hero)

        if n_event == "Apply":

            # Checks if the inputs are valid if they are it will return the updated metadata
            if not n_values["-MANA-"].replace("%", "").isdigit():
                sg.popup("Your mana input is invalid! it has to be a number")
                continue
            elif not n_values["-LEVEL-"].isdigit():
                sg.popup("Your level input is invalid! it has to be a number")
                continue
            elif int(n_values["-MANA-"].replace("%", "")) > 100 or int(n_values["-MANA-"].replace("%", "")) < 0:
                sg.popup("Your mana input is invalid! it can only be between 0 and 100")
                continue
            elif int(n_values["-LEVEL-"].replace("%", "")) > 3 or int(n_values["-LEVEL-"].replace("%", "")) < 1:
                sg.popup("Your level input is invalid! it can only be between 1 and 3")
                continue
            else:
                window.close()
                lvl = n_values["-LEVEL-"]
                return {"hero": selected_hero, "mana": n_values["-MANA-"].replace("%", ""),
                        "level": 1 if lvl == 1 else 1.25 if lvl == 2 else 1.5}


def export_as_overwatch_code(close_window):
    """
    Will convert the board configuration to overwatch code that can be pasted as a rule

    :param close_window: The window to close after exported
    :return: None
    """

    # The rule template
    rest = """
    rule("ChessWatch Match - By Darkshadow - [github link]")
    {
        event
        {
            Ongoing - Global;
        }
        
        conditions
        {
        Count Of(All Players(All Teams)) == 1;
        }
    """

    actions = "\n        actions\n        {\n"

    # Loops through every button/hero on the grid
    for i in range((columns * rows) * 2):

        # Gets the `i` button and hero name
        key = f"-T{1 if i <= columns * rows - 1 else 2}_{i}-"
        elem = main_window[key]
        hero_name = elem.metadata['hero'].lower()

        # Changes some names according to specific rules
        if hero_name == "torbjorn":
            hero_name = "torbjörn"
        if hero_name == "lucio":
            hero_name = "lúcio"
        if hero_name == "soldier76":
            hero_name = "soldier: 76"

        # Adds every action according to that hero's metadata
        if elem.metadata['hero']:
            actions += f"           \"{hero_name.title()} - ChessWatch Match Generator\"\n"
            actions += f"           Create Dummy Bot(Hero({hero_name.title()}), Team {1 if i <= columns * rows - 1 else 2}, -1, {positions[i]}, Vector(0, 0, 0));\n"
            actions += f"           Global.last_created_entity = Last Created Entity;\n"
            actions += f"           Global.last_created_entity.mana = {elem.metadata['mana']};\n"
            actions += f"           Global.last_created_entity.level = {elem.metadata['level']};\n\n"
    actions += "        }"

    # Removes the first tabs (to prettify the code)
    rest += actions + "\n    }"
    rest = rest.splitlines(True)
    for count, line in enumerate(rest):
        rest[count] = line.replace("    ", "", 1)

    # Asks the user to save the exported code into a file
    file = sg.popup_get_file("Save As Export...", save_as=True, no_window=True,
                             file_types=(("txt", "*.txt"),), default_extension=".txt")

    # Writes the code into that file
    if file:
        with open(file, "w") as f:
            f.writelines(rest[1:])
            f.close()

        # Closes the export settings window and then opens the saved file
        close_window.close()
        os.startfile(file)


def export(close_window):
    """
    Exports the board configuration to an encoded document

    :param close_window: The window to close after exported
    :return: None
    """

    # Puts every cell metadata into a dictionary
    settings = {}

    # Loops through every cell on the board
    for i in range((columns * rows) * 2):
        key = f"-T{1 if i <= columns * rows - 1 else 2}_{i}-"
        settings[key] = {
            "hero": main_window[key].metadata["hero"],
            "mana": main_window[key].metadata["mana"],
            "level": main_window[key].metadata["level"]
        }

    # Converts the dict to str
    settings = f"{settings}"

    # Encodes the data with `base64`
    message_bytes = settings.encode('ascii')
    base64_bytes = base64.b64encode(message_bytes)
    base64_message = base64_bytes.decode('ascii')

    # Pops up a save as prompt window
    file = sg.popup_get_file("Save As Export...", save_as=True, no_window=True,
                             file_types=(("mgconfig", "*.mgconfig"),), default_extension=".mgconfig")
    if file:
        with open(file, "w") as f:
            f.writelines(base64_message)
        close_window.close()


def import_config(clicked_file=""):
    """
    Imports an encoded document and decodes it and then applies the settings to the board

    :param clicked_file: if a config file was double clicked it would be it's path
    :return: None
    """

    # Sets the file path
    if not clicked_file:
        file = sg.popup_get_file("Load...", no_window=True, file_types=(("mgconfig", "*.mgconfig"),),
                                 default_extension=".mgconfig")
    else:
        file = clicked_file

    if file:

        # Decodes the base64 data
        settings = ast.literal_eval(base64.b64decode(open(file).read()).decode("utf-8"))

        # Loops through every key (cell key) in the data and updates it by the data inside of it
        for key in settings.keys():
            update_cell_image(key, settings[key]["hero"])
            main_window[key].metadata["mana"] = settings[key]["mana"]
            main_window[key].metadata["level"] = settings[key]["level"]


def randomize_deck():
    """
    Randomizes the board, with random heroes in random positions
    
    :return: None
    """

    max_per_team = 5
    hero_counter = 0

    # Clears the whole board
    for i in range((columns * rows) * 2):
        update_cell_image(f"-T{1 if i <= columns * rows - 1 else 2}_{i}-")

    # Loops through every cell on team 1's grid and checks if the hero count haven't exceeded the limit if not it will
    # randomly pick if to place a hero there or not if it picks to place a hero it will update that cell with the image
    # of a random hero from the `ow_heroes` list
    for team1_counter in range((columns * rows) - 1):
        if hero_counter < max_per_team:
            key = f"-T1_{team1_counter}-"
            if random.randint(1, 5) == 1:
                hero_counter += 1
                update_cell_image(key, random.choice(ow_heroes))

    # If there are less then 2 heroes (0/1) it will place 2 more random heroes in random positions
    if hero_counter < 2:
        for i in range(2):
            random_hero = random.choice(ow_heroes)
            update_cell_image(f"-T1_{random.randint(0, (columns * rows) - 1)}-", random_hero)

    # Loops through every cell on team 2's grid and checks if the hero count haven't exceeded the limit if not it will
    # randomly pick if to place a hero there or not if it picks to place a hero it will update that cell with the image
    # of a random hero from the `ow_heroes` list
    hero_counter = 0
    for team2_counter in range((columns * rows), columns * rows * 2 - 1):
        if hero_counter < max_per_team:
            key = f"-T2_{team2_counter}-"
            if random.randint(1, 5) == 1:
                hero_counter += 1
                update_cell_image(key, random.choice(ow_heroes))

    # If there are less then 2 heroes (0/1) it will place 2 more random heroes in random positions
    if hero_counter < 2:
        for i in range(2):
            random_hero = random.choice(ow_heroes)
            update_cell_image(f"-T2_{random.randint(columns * rows, (columns * rows) * 2 - 1)}-", random_hero)


# If a config file was opened it will click the `import` button
if len(sys.argv) == 2:
    main_window["-IMPORT-"].click()

# Applies the right click event to every button in the grid
for i in range((columns * rows) * 2):
    main_window[f"-T{1 if i <= columns * rows - 1 else 2}_{i}-"].bind('<Button-3>', ' +RIGHT CLICK+')

# ----- MAIN PROGRAM LOOP ----- #

while True:

    event, values = main_window.read()

    if event == sg.WINDOW_CLOSED:
        break

    # If a button was right clicked it will set the var `last_right_clicked_btn` to that button key
    if "+" in event:
        if "select" not in event.lower() and "clea" not in event.lower():
            last_right_clicked_btn = event.replace(" +RIGHT CLICK+", "")

    # If the user clicked the clear deck button it will loop through the whole board and update the cell's metadata
    # to the default one
    if event == "-CLEAR_DECK-":
        for i in range((columns * rows) * 2):
            key = f"-T{1 if i <= columns * rows - 1 else 2}_{i}-"
            update_cell_image(key)
            main_window[key].metadata["level"] = 1
            main_window[key].metadata["mana"] = "0"

    # If the user has clicked the randomize deck button it will call the randomize deck function
    if event == "-RANDOMIZE_DECK-":
        randomize_deck()

    # If the user has clicked the select hero button on the right click menu of a cell it will call the
    # `create_hero_selection_window` which creates a hero selection window, after the user has selected a hero it will
    # update the last right clicked button's image and metadata to that selected hero
    if event == "Select hero":
        hero = create_hero_selection_window()
        if hero:
            update_cell_image(last_right_clicked_btn, hero)

    # If the user has clicked edit properties in the right click menu it will call the function create_properties_editor
    # _window and pass it the hero's metadata incase the hero was edited before and then after the user applies the new
    # metadata it will update the hero's metadata to the updated metadata
    if event == "Edit properties":
        if main_window[last_right_clicked_btn].metadata["hero"]:
            new_metadata = create_properties_editor_window(main_window[last_right_clicked_btn].metadata["hero"],
                                                           last_right_clicked_btn,
                                                           main_window[last_right_clicked_btn].metadata)
            if new_metadata:
                main_window[last_right_clicked_btn].metadata = new_metadata
        else:
            sg.Popup("This cell is empty, You have to have a hero there to edit it's properties", title="Error")

    # If the user has clicked the clear button in the right click menu it will update the cell's image
    # and it's metadata to default
    if event == "Clear":
        update_cell_image(last_right_clicked_btn)
        main_window[last_right_clicked_btn].metadata["level"] = 1
        main_window[last_right_clicked_btn].metadata["mana"] = "0"

    # If the user has clicked the export button it will call the create_export_settings_window function which
    # pops up the export settings window
    if event == "-EXPORT-":
        create_export_settings_window()

    # If the user has clicked the import button it will call the import_config function which
    # updates all of the cells according to the imported data
    if event == "-IMPORT-":
        import_config("" if len(sys.argv) < 2 else sys.argv[1])

main_window.close()
