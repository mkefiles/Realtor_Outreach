# Realtor Outreach

This is really a two-part (as detailed below).

## Part One: Scrape the Realtor Database

This part used Python and Selenium to scrape what ended up being over 33,000 realtors' contact information. Naturally, the data has since been cleaned to remove any bad data-points.

## Part Two: Automate Dynamic Emails

This part automated the emailing process, with dynamic emails imported from a JSON file to off-set SPAM filters, across five email accounts using Pythong (SSL / win32com). The *dynamic* aspect is a bit more... sophisticated than it sounds. I had Python import a JSON file with multiple variants of each section of the overall email and it would, by using Regular Expressions, swap out "placeholders" from the variants. Then it would feed the email and, by use of a randomized process, would send it via one of five separate emails (i.e., two iCloud emails, two GMail emails and one Outlook email). iCloud and Gmail required slightly different code to send. Outlook required installation of Microsoft Outlook Classic and I had Python take over control of Outlook because, otherwise, Microsoft makes it difficult to send via Python (I could not locate the ability to do it).
