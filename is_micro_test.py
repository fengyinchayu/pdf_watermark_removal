def is_micro_test(name):
    keywords = [
        "Total Aerobic",
        "Mold",
        "Yeast",
        "Coliform",
        "Salmonella",
        "E.coli",
        "E. coli",
        "Staphylococcus"
    ]
    return any(k.lower() in name.lower() for k in keywords)
