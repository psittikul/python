# Make all necessary imports
from selenium import webdriver

# Launch Firefox and open the Orio OSC portal
browser = webdriver.Firefox()
browser.get('http://osc.orioautoparts.com/')

# Log in
userElem = browser.find_element_by_id('topname')
userElem.clear()
userElem.send_keys('patricia')
pwElem = browser.find_element_by_id('toppassword')
pwElem.clear()
pwElem.send_keys('SittikulOAP2016')
pwElem.submit()

# Click "View As Company" to begin going through the list of OSCs
browser.find_element_by_partial_link_text('View as Company').click()

# You only want OSCs, so only the shops with the type "OSC" should be
# queued for processing
allOSC = browser.find_elements_by_css_selector('[data-site="OSC"]')

# Now that you have the OSCs identified, go through each one
for osc in allOSC:
    # View the order history of each OSC (you need to know the OSC's
    # name to make sure you click the right order history button)
    
    # ------ Check this piece of code to make sure it works ------
    oscName = str(osc.text)
    oscName = oscName.split("\n")
    oscName = oscName[1].split(" ")
    for num in range(0, int(len(oscName) - 3)):
        if num != 4:
            name += oscName[num] + " "
        else name+=oscName[num]
    pathVar = '//*[@data-name="'+str(name)+'"]/span[4]/i'
    browser.find_element_by_xpath(pathVar).click()
    # ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    """ Look for the amount of money in their Co-Op Accrual
        1. Make an array of all tables on the page
        2. The very last table (at index len(tables) - 1) contains the
        information we need
        3. Get the text of that table and turn it into a string for parsing
        4. Parse the text of the table to just get the co-op accrual value
        5. Store the OSC's name and co-op funds in an Excel spreadsheet
    """
    tables = browser.find_elements_by_tag_name('table')
    index = len(tables) - 1
    containsAnswer = str(tables[index].text)
    containsAnswer = containsAnswer.split("\n")
    tableRowAnswer = containsAnswer[1].split(" ")
    coopFunds = tableRowAnswer[9]
    
    
    
    
