from modules.gui import GUI
from modules.search_logic import SearchLogic

if __name__ == "__main__":
    gui_instance = GUI(search_logic_instance=None)

    # Passando a instância de GUI corretamente
    search_logic_instance = SearchLogic(gui_instance=gui_instance)

    gui_instance.search_logic_instance = search_logic_instance

    gui_instance.create_gui()
