SELECT
    initiativecampaign,
    retail_week,
    ad_type_tactic,
    update_total_budget,
    platform,
    publisher,
    start_date,
    end_date
FROM nike_na.copy_of_budgets_and_goals_fy26_pld_1772748449255
WHERE
    budget_source = 'NDDC - Display'
    AND pmg_channel_team = 'Programmatic'
    AND CURRENT_DATE >= start_date
    AND CURRENT_DATE <= end_date
ORDER BY retail_week DESC