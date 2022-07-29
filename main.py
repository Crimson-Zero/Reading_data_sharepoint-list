from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.listitems.listitem import ListItem
from bs4 import BeautifulSoup
import matplotlib.pyplot as plt

site_url = "https://[your tenant name].sharepoint.com/sites/{sitename}"

username = "username"
password = "password"


list_data = []
ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
target_list = ctx.web.lists.get_by_title("list name")

VPN_array = []
internet_array = []
outlook_array = []
teams_array = []


vpn_text = "VPN"
internet_text= "internet"
outlook = "outlook"
teams = "teams"

issue_list = [vpn_text,internet_text,outlook,teams]

paged_items = target_list.items.paged(50).get().execute_query()
for index,item in enumerate(paged_items):
    
    get_body = item.properties["item name"]
    
    if get_body is not None:
        
        soup = BeautifulSoup(get_body)
        text_out  = soup.get_text()
        list_data.append(text_out)
        
        if (vpn_text) in text_out:
            VPN_array.append(text_out)
            
        if outlook in text_out:
            outlook_array.append(text_out)
        
        if (internet_text or "Internet") in text_out:
            internet_array.append(text_out)
        
        if teams in text_out:
            teams_array.append(text_out)
            

print(list_data)

vpn_count = len(VPN_array)
outlook_count = len(outlook_array)
internet_count = len(internet_array)
team_count = len(teams_array)

y_axis = [vpn_count,outlook_count,internet_count,team_count]
plt.bar(issue_list, y_axis, color ='maroon',
        width = 0.4)
 
plt.xlabel("Reported Issues")
plt.ylabel("No of issues ")
plt.title("IT Tickets")
plt.show()
