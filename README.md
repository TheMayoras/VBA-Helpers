# VBA-Helpers
These are some homemade classes and modules that I have made and re-used in all my large VBA projects.  Most of these classes are used in Access, but I do not see any reason why they can't be used in other MS Office applications.

# CEventBoxManagerForm/CEventBoxManagerReport
Use these classes when you want to gather data from a different form without having to use global variables and the like.  I use these classes more than any other.  The most common workflow for these classes is to open a open the form in a button's `OnClick` event. Then when the form is closed, you can gather the data from the form by using the `BeforeFormClose` event. These classes do a lot of the dirty work for me.  

# PublishVersion
I use this module to publish my `.accdb` files as `.accde`.  It also allows me to make every form a `pop-up` (which I prefer), and add a `Me.Move 0, 0` to the load event forms and reports.  Adding `Me.Move 0, 0` prevents them from disappearing when they are opened.
