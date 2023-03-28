if template.name == 'collection' and len(collection.all_tags) > 0:
    categories = []
    for tag in collection.all_tags:
        if "_" in tag:
            category = tag.split("_")[0]
            categories.append("|" + category)
    base_cat_array = list(set(categories[1:]))
    
    custom_ordered_categories = section.settings.custom_ordered_categories.split(',')
    custom_categories = []
    for custom_ordered_category in custom_ordered_categories:
        custom_category = custom_ordered_category.strip()
        if custom_category not in base_cat_array:
            continue
        custom_categories.append("|" + custom_category)
        
    cat_array = list(set(custom_categories[1:] + base_cat_array))
