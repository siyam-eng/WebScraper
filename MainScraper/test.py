from get_codes import find_indices
import re

def find_code(code):
    """Finds all codes matching with the given code"""
    # codes = soup.findAll(text=re.compile(f'{code}'))
    codes = ['LUE LISÄÄ MULTI-TABS -TUOTTEISTA', ' PM-FI-MUL-21-00003']
    codes2 = []
    for c in codes:
        code_list = re.findall(f'[A-Z-/]*{code}[A-Z0-9a-z-/]*', str(c))
        for code2 in code_list:
            code2 = code2.lstrip('-') if code2.startswith('-') else code2
            code2 = code2.lstrip('/') if code2.startswith('/') else code2
            if '/' in code2 or '-' in code2:
                # print(code2)
                # limiting the lenth of the code upto 3rd seperator
                if code2.count('-') > 4:
                    third_hyphen = find_indices(code2, '-')[4]
                    code2 = code2[: third_hyphen]
                if code2.count('/') > 3:
                    third_slash = find_indices(code2, '/')[3]
                    code2 = code2[: third_slash]
                if code2.count('/') > 2 and '-' in code2:
                    code2 = code2[: code2.index('-')]

                if code2.count('/') > 2 and code.count('-') > 4:
                    # codes2.append(code2)
                    print(code2)
                codes2.append(code2)
    return codes2


print(find_code('MUL'))