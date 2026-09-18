import re

with open('database.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Locate the CREATE TABLE monthly_expenses block
create_block_pattern = re.compile(
    r"(    # Tabel Monthly Expenses \(Cheltuieli lunare\)\n    cursor\.execute\('''\n        CREATE TABLE IF NOT EXISTS monthly_expenses \([\s\S]*?        \)\n    '''\)\n)",
    re.MULTILINE
)

match_create = create_block_pattern.search(content)
if match_create:
    create_block = match_create.group(1)
    
    # Remove the create block from its original location
    content = content.replace(create_block, '')
    
    # Locate the ALTER block for monthly_expenses
    alter_block_start = "    # Migrare automată pentru monthly_expenses"
    alter_idx = content.find(alter_block_start)
    
    if alter_idx != -1:
        # Insert the create block before the alter block
        content = content[:alter_idx] + create_block + "\n" + content[alter_idx:]
        
        with open('database.py', 'w', encoding='utf-8') as f:
            f.write(content)
        print("Successfully rearranged monthly_expenses initialization.")
    else:
        print("Could not find alter block.")
else:
    print("Could not find create block.")
