
![Logo](https://snipeitapp.com/img/logos/snipe-it-logo-xs.png)


# SnipeIT Updater


This script was created for my old workplace, where we had to auto update and auto add new computer data to the SnipeIT database ( https://snipeitapp.com/ )

I never finished the project, so don't expect a finish and ready to use project. Some features were removed due to them being company specific.
## What can it add to SnipeIT?

- location_id: Where the computer is located.
- asset_tag: Naming of computer, example: PC-123456
- model_id: The model of the computer.
- serial: The serial number of computer.
- ram (custom field): The ram amount of computer in GB.
- cpu (custom field): The CPU name of the computer.
- storage (custom field): The amount of storage converted to GB. (128, 256, 512 etc.)


## FAQ

#### How does It work?

The script grabs data both from SnipeIT database and gathers the computer data. It will use these data to insert into the SnipeIT database.

#### How do I use it?

You have to start by adjusting the top variables to your SnipeIT site specific data. Both the key from panel and the API url to your site.


#### Will the script recieve updates?

I'm not planning on updating it any further. It's an older script that I wanted to share.
