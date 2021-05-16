"""
"Magic Session Mixer" is a gui tool for live audio control
and viewing of active audio sessions.

Features
--------
    List all active sessions, updates via callbacks.
    Change volume or mute via gui.
    When changes to the session occur, the gui will update
      and display session state, volume, and mute accordingly.
    In total:
        Everyting you can do with the build-in Windows audio mixer:
          Minus: Speaker volume support and live audio stream visualisation
          Plus: more accurate/ faster representation of the audio session state
                and presence.

Note
----
pycaw.magic needs to be imported before any other pycaw or comtypes import.
"""

import os
from contextlib import suppress
# _____ ADVANCED WINDOWS COMMUNICATION _____
from ctypes import windll
# ____________ GUI WITH TKINTER ____________
from tkinter import BOTTOM, HORIZONTAL, DoubleVar, StringVar, Tk
from tkinter.ttk import Button, Frame, Label, Scale, Separator, Style

from pycaw.magic import MagicManager, MagicSession  # isort:skip
from pycaw.constants import AudioSessionState  # isort:skip

# _________ GUI SCALE ON HIGH DPI _________
windll.shcore.SetProcessDpiAwareness(2)
# SetProcessDpiAwareness = 2: Per monitor DPI aware.
# shcore checks for the DPI when the app is
# started and adjusts the scale factor whenever the DPI changes.
# Normally tkinter apps are not automatically scaled.


class RootFrame(Tk):
    """Mixer Window"""
    def __init__(self):
        super().__init__()
        # __________ ROOT WINDOW MENU BAR __________
        self.geometry("500x600")
        self.title("Magic Session Mixer")
        # set icon:
        dirname = os.path.dirname(__file__)
        icon_path = os.path.join(dirname, 'magic_session_mixer.ico')
        self.iconbitmap(icon_path)

        # ________________ STYLING ________________
        s = Style()

        s.configure("Active.TFrame", background='#007AD9')
        # s.configure("Inactive.TFrame", background='#A2A2A2')
        s.configure("Inactive.TFrame", background='#AF3118')

        s.configure("Unmuted.TButton", foreground='#007AD9')
        # s.configure("Muted.TButton", foreground='#C33C54')
        s.configure("Muted.TButton", foreground='#AF3118')

        # _________________ HEADER _________________
        header = Frame(self)
        title = Label(header, text="Magic Session Mixer",
                      font=("Calibri", 20, "bold"))

        # ________________ SEPARATE ________________
        separate = Separator(header, orient=HORIZONTAL)

        # ____________ ARRANGE ELEMENTS ____________
        header.columnconfigure(0, weight=1)
        title.grid(row=0, column=0, pady=10, sticky="W")
        separate.grid(row=1, column=0, sticky="EW", pady=10)

        # _______________ PACK FRAME _______________
        header.pack(fill='x', padx=15)

        # ____________ BUILD WITH PYCAW ____________
        turbo_anonym = Label(self, text="created by TurboAnonym with pycaw",
                             font=("Consolas", 8, "italic"))
        turbo_anonym.pack(side=BOTTOM, fill="x", pady=6, padx=6)


class AppRow(MagicSession, Frame):
    """
    Row for each app in mixer.
    handles refreshing the gui if session is changed external.
    handles user input and changing session volume/mute.
    """

    def __init__(self, root_frame_instance):
        super().__init__(volume_callback=self.update_volume,
                         mute_callback=self.update_mute,
                         state_callback=self.update_state)

        self.root_frame_instance = root_frame_instance

        # ______________ DISPLAY NAME ______________
        self.app_name = self.magic_root_session.app_exec

        print(f":: new session: {self.app_name}")
        # ______________ CREATE FRAME ______________
        # super(MagicSession, self).__init__(root_frame_instance)
        Frame.__init__(self, root_frame_instance)

        # _______________ NAME LABEL _______________
        self.name_label = Label(self,
                                text=self.app_name,
                                font=("Consolas", 12, "italic"))

        # _____________ VOLUME SLIDER _____________
        self.volume_slider_state = DoubleVar()

        self.volume_slider = Scale(self,
                                   variable=self.volume_slider_state,
                                   command=self._slide_volume,
                                   from_=0, to=100,
                                   takefocus=False,
                                   orient=HORIZONTAL)

        # set initial:
        self.volume_slider_state.set(self.volume * 100)

        # ______________ MUTE BUTTON ______________

        self.mute_button_state = StringVar()

        self.mute_button = Button(self,
                                  style="",
                                  textvariable=self.mute_button_state,
                                  command=self._toogle_mute,
                                  takefocus=False)

        # set initial:
        self.update_mute(self.mute)

        # _____________ SESSION STATUS _____________
        self.status_line = Frame(self, style="", width=6)

        # set initial:
        self.update_state(self.state)

        # ________________ SEPARATE ________________
        self.separate = Separator(self, orient=HORIZONTAL)

        # ____________ ARRANGE ELEMENTS ____________
        # set column[1] to take the most space
        # and make all others as small as possible:
        self.columnconfigure(1, weight=1)

        # grid
        self.name_label.grid(row=0, column=0, columnspan=2, sticky="EW")
        self.mute_button.grid(row=1, column=0)
        self.volume_slider.grid(row=1, column=1, sticky="EW", pady=10, padx=20)
        self.separate.grid(row=2, column=0, columnspan=3, sticky="EW", pady=10)
        self.status_line.grid(row=0, rowspan=2, column=2, sticky="NS")

        # _____________ DISPLAY FRAME _____________
        self.pack(pady=0, padx=15, fill='x')

    def update_volume(self, new_volume):
        """
        when volume is changed externally
        (see callback -> AudioSessionEvents -> OnSimpleVolumeChanged )
        """
        # compare if the windows callback is because we set the slider:
        # if so drop update since we made the change
        # and all is already up to date - this will prevent lagg
        print(f"{self.app_name} volume: {new_volume}")
        self.volume_slider_state.set(new_volume*100)

    def update_mute(self, new_mute):
        """ when mute state is changed by user or through other app """
        if new_mute:
            icon = "🔈"
            self.mute_button.configure(style="Muted.TButton")
        else:
            icon = "🔊"
            self.mute_button.configure(style="Unmuted.TButton")

        # .set is a method of tkinters variables
        # it will change the button text
        print(f"{self.app_name} mute: {icon}")
        self.mute_button_state.set(icon)

    def update_state(self, new_state):
        """
        when status changed
        (see callback -> AudioSessionEvents -> OnStateChanged)
        """
        print(f"{self.app_name} state: {new_state}")
        if new_state == AudioSessionState.Inactive:
            # AudioSessionStateInactive
            self.status_line.configure(style="Inactive.TFrame")

        elif new_state == AudioSessionState.Active:
            # AudioSessionStateActive
            self.status_line.configure(style="Active.TFrame")

        elif new_state == AudioSessionState.Expired:
            # AudioSessionStateExpired
            self.status_line.configure(style="TFrame")
            """when session expires"""
            print(f":: closed session: {self.app_name}")
            self.destroy()

    def _slide_volume(self, value):
        """ when slider moved by user """
        new_volume = float(value)/100
        # check if new user value really is new: (ttk bug)
        if self.volume != new_volume:
            # since self.volume is true data through windows
            # it will generally differ, but 1.0 == 1
            print(f"with pycaw: {self.app_name} volume: {new_volume}")
            self.volume = new_volume

    def _toogle_mute(self):
        """ when mute button pressed """
        new_mute = self.toggle_mute()

        self.update_mute(new_mute)


def main():
    # ___________ CREATE ROOT WINDOW ___________
    root_frame_instance = RootFrame()

    # ________ START THE MAGIC ________
    # args and kwargs are passed
    MagicManager.magic_session(AppRow, root_frame_instance)

    # ______________ START WINDOW ______________
    with suppress(KeyboardInterrupt):
        root_frame_instance.mainloop()

    print("\nTschüss")


if __name__ == '__main__':
    main()
