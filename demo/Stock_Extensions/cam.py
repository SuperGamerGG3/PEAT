"""
PEAT EXTENSION
author = Beffy
name = Cams
filename = cam
version = 1.0

requirements:
- None
"""

EXT_NAMESPACE = "cam"

help_dict = {
    "open": "Opens camera veiw of specified cam (arg1)",
}

def load_extension():
    peat.register_command("open", cmd_cam) # type: ignore


    peat.register_help(help_dict) # type: ignore

def cmd_cam(a1, a2, title):

    if not a1:
        peat.print("Expected a camera index.") # type: ignore
        return

    try:
        peat.set_camera_index(int(a1)) # type: ignore
    except ValueError:
        peat.print("Camera index must be a number.") # type: ignore
        return

    peat.set_camera_mode(True) # type: ignore
    peat.print(f"Opening camera index {a1}.") # type: ignore