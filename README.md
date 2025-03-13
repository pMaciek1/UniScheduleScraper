# University schedule scraper
Simple program I built to scrape the webpage of my university for the schedules.
It searches for div with schedule for my course, downloads all the pdfs, and formats them to only save the pages with my classes.
Additionally it generates excel file, with classes for my specialization.

## Requirements
- python 3.12+
- requests (accessing the webpage)
- bs4 (parsing the html)
- os (operating on files)
- shutil (operating on files)
- urllib (formatting the URL string)
- pypdf (operating on .pdf files)
- xlsxwriter (operating on .xlsx files)
