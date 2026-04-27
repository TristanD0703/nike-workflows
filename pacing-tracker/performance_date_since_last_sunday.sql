SELECT 
    campaign_name, 
    harmonized_publisher_platform, 
    harmonized_partner, harmonized_budget_type, 
    SUM(impressions) as impressions, 
    SUM(clicks) as clicks, 
    SUM(media_spend) as spend, 
    harmonized_campaign_initiative, 
    pmg_channel,
    harmonized_sub_campaign
FROM nike_na_custom.analytics_omnichannel 
WHERE harmonized_budget_type = 'nddc - display' 
AND media_spend > 0
AND pmg_channel = 'display'
AND date >= (current_date - extract(dow from current_date)::integer)
AND date < CURRENT_DATE
GROUP BY harmonized_sub_campaign, campaign_name, harmonized_publisher_platform, harmonized_partner, harmonized_budget_type, harmonized_campaign_initiative, pmg_channel