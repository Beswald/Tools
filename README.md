# Tools
Collection of utilities and tools

LuaMacros
-------------------------
By using this application and a custom lua script, I created a set of macros that can be launched using a generic (i.e. cheap) usb keypad. I'm only including the scripts I created, you must download LuaMacros from the source below.

Using this method I invoke programs, simulate key combinations and modify system settings. Used mostly for shortcuts while doing development including several keys that are context aware. Using one key I can lookup all references in visual studio and perform a generic find (Ctrl + f) in most other places.

- MacroPad.lua - My custom script that takes keypresses, converts them to key combinations, launches programs and runs other scripts.

- KeyLayout.xlsx - Key map for my keypad, this will be different for different keypads. The correspoding key id's will have to change accordingly.

- Scripts folder - Contains various visual basic scripts launched by macro keys. Usually multi step operations or actions that needed file access etc...

Usage:
    LuaMacros.exe -r MacroPad.lua


LuaMacros resources:
MessageBoard/Help/Faq: http://www.hidmacros.eu/forum/viewforum.php?f=9
Download and Version Information:  http://www.hidmacros.eu/forum/viewtopic.php?f=10&t=241
