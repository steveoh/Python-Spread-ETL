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

zip_template = Template("""private static Dictionary<string, List<GridLinkable>> CacheZipCodes() 
{ 
    var gridLookup = new List<ZipGridLink> { {% for item in items %}
        new ZipGridLink({{item[0]}}, "{{item[1]}}", {{item[2]}}){% if not loop.last %},{% endif %}{% endfor %}
    }; 
        
    return BuildGridLinkableLookup(gridLookup);
}""")

places_template = Template("""\n\nprivate static Dictionary<string, List<GridLinkable>> CachePlaces()
{
    var gridLookup = new List<PlaceGridLink> { {% for item in items %} 
        new PlaceGridLink("{{item[0]}}", "{{item[1]}}", {{item[2]}}){% if not loop.last %},{% endif %}{% endfor %} 
    };
    
    return BuildGridLinkableLookup(gridLookup);
}""")

usps_template = Template("""\n\nprivate static Dictionary<string, List<GridLinkable>> CacheDeliveryPoints()
{
    var gridLookup = new List<UspsDeliveryPointLink> { {% for item in items %} 
        new UspsDeliveryPointLink({{item[0]}}, "{{item[1]}}", 0, "{{item[2]}} {{item[3]}}", {{item[6]}}, {{item[7]}}){% if not loop.last %}, {% endif %}{% endfor %} 
    };
    
    return BuildGridLinkableLookup(gridLookup);
}""")


with open("both.txt", "w") as text_file:
    text_file.write(zip_template.render(items = get_list_from("ZipCodesInAddressQuadrants")))
    text_file.write(places_template.render(items = get_list_from("Cities/Placenames/Abbreviations w/Address System")))
    text_file.write(usps_template.render(items = get_list_from("USPS Delivery Points")))

print 'updated places list.'

os.startfile("both.txt")

