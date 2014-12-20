# OpenFolder

[AutoIt](https://www.autoitscript.com/site/) script to quickly search subfolders of a given folder by substring and open webapps.

OpenFolder is created to quickly open a specific business folder in a list of (thousands) of folders with customer names like `Customer Name 000000`. A real time saver, especially with a lot of similar customer names like municipalities or hospitals.

## Setup

Store `OpenFolder.exe` somewhere on the network where users can access it. Create a convenient shortcut for the end user, for instance on their desktop or task bar. The shortcut must provide the executable with a path to the main folder.

If your business folders contain an identifier used in webapps, you can configure the right mouse button via `OpenFolder.ini` - an example is provided.

## Daily usage

The user is presented with a search box and an empty list of foldernames. The user can type a part of a foldername to list all folders that match this substring. The user can then double click to open a folder - or use the right mouse button for extra options like copy an identifier or open a configured webapp.