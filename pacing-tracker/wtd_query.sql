WITH PME as (
    SELECT 
        LOWER(harmonized_publisher_platform) as harmonized_publisher_platform, 
        LOWER(harmonized_partner) as harmonized_partner, 
        harmonized_budget_type, 
        SUM(impressions) as impressions, 
        SUM(clicks) as clicks, 
        SUM(media_spend) as spend, 
        (harmonized_campaign_initiative || ' ' || harmonized_sub_campaign) as big_boi_campaign, 
        pmg_channel,
        ad_type_tactic,
        retail_week
    FROM nike_na_custom.analytics_omnichannel 
    WHERE harmonized_budget_type = 'nddc - display' 
    AND media_spend > 0
    AND pmg_channel = 'display'
    AND date >= (current_date - extract(dow from current_date)::integer)
    AND date < CURRENT_DATE
    GROUP BY 
        harmonized_sub_campaign, 
        harmonized_publisher_platform, 
        harmonized_partner, 
        harmonized_budget_type, 
        harmonized_campaign_initiative, 
        pmg_channel,
        ad_type_tactic,
        retail_week
),
PLD AS (
    SELECT
    LOWER(initiativecampaign) as initiativecampaign,
    retail_week,
    LOWER(ad_type_tactic) as tactic,
    update_total_budget,
    LOWER(platform) as platform,
    LOWER(publisher) as publisher,
    start_date,
    end_date
FROM nike_na.copy_of_budgets_and_goals_fy26_pld_1772748449255
WHERE
    budget_source = 'NDDC - Display'
    AND pmg_channel_team = 'Programmatic'
    AND CURRENT_DATE >= start_date
    AND CURRENT_DATE <= end_date
ORDER BY retail_week DESC
)
SELECT * 
FROM PLD FULL OUTER JOIN PME ON 
(platform = harmonized_partner OR platform = 'google ads')
AND publisher = harmonized_publisher_platform
AND PLD.retail_week = PME.retail_week
AND big_boi_campaign LIKE '%' || initiativecampaign || '%'


