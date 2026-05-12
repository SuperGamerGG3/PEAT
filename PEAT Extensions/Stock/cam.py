# Stock extension, created by Beffy
# Requires a camera

EXT_NAMESPACE = "cam"

help_dict = {
    "open": "Opens camera veiw of specified cam (arg1)",
}

def load_extension():
    peat.register_command(EXT_NAMESPACE, "open", cmd_cam) # type: ignore


    peat.register_help(EXT_NAMESPACE, help_dict) # type: ignore

def cmd_cam(a1, a2, title):

    if not a1:
        peat.voice_print("Expected a camera index.") # type: ignore
        return

    try:
        peat.set_camera_index(int(a1)) # type: ignore
    except ValueError:
        voice_print("Camera index must be a number.") # type: ignore
        return

    peat.set_camera_mode(True) # type: ignore
    peat.voice_print(f"Opening camera index {a1}.") # type: ignore