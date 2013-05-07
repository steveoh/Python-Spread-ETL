import gspread, os, pickle, ConfigParser
from jinja2 import Template

def get_credentials():
    config = ConfigParser.RawConfigParser()
    
    secrets = 'secrets.cfg'
    
    config.read(secrets)
    
    user = config.get("Google", "username")
    password = config.get("Google", "password")
    
    return user, password

def get_list_from(spreadsheet):
    print 'getting data from {}'.format(spreadsheet) 
    
    gc = None
    
    try:
        gc = pickle.load(open('conf.p', 'rb'))
    except: 
        pass
    
    if gc is None:    
        creds = get_credentials() 
        gc = gspread.login(creds[0], creds[1])
        pickle.dump(gc,open('conf.p','wb'))
    
    # Open a worksheet from spreadsheet with one shot
    wks = gc.open(spreadsheet).sheet1
    
    list_of_lists = wks.get_all_values()
    list_of_lists.pop(0)
    
    return list_of_lists

zip_template = Template("""private static Dictionary<string, List<GridLinkable>> GetZipCodeToGridLookup() 
{ 
    var gridLookup = new List<ZipGridLookup> { 
        {% for item in list %} new ZipGridLookup({{item[0]}}, "{{item[1]}}", {{item[2]}}){% if not list.last %},{% endif %}
        {% endfor %}}; 
    return BuildGridLinkableLookup(gridLookup);
'}""")

places_template = Template("""private static Dictionary<string, List<GridLinkable>> GetPlaceToGridLookup()
{
    var gridLookup = new List<PlaceGridLookup> {
        {% for item in list %}
        new PlaceGridLookup("{{item[0]}}", "{{item[1]}}", {{item[2]}})
        {% if not list.last %}, {% endif %}
        {% endfor %} };
    
    return BuildGridLinkableLookup(gridLookup);
}""")

with open("both.txt", "w") as text_file:
    text_file.write(zip_template.render(list = get_list_from("ZipCodesInAddressQuadrants")))

print 'updated zip file list.'
    
with open("both.txt", "a") as text_file:
    text_file.write(places_template.render(list = get_list_from("Cities/Placenames/Abbreviations w/Address System")))

print 'updated places list.'

os.startfile("both.txt")

