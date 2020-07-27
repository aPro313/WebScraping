
Scraping colleges from https://www.niche.com/colleges/search... 
and the issue is that after a few requests server bans the scraper with 403 response even though robots.txt doesn't seem to contain any crawling restrictions. 
First I've implemented the scraper due to the requisites but when it came to actual pagination crawling it turned out that one can't scrape more then 10 pages.
So the first solution was to provide a proxy rotation using free proxies, but even though it worked for a while still the result wasn't obtained, so eventually I've used crawlera api (payed proxy rotation) to scrape all the data and it worked like a charm.

