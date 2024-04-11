from difflib import SequenceMatcher

def compare_strings(s1, s2):
    s = SequenceMatcher(lambda x: x in ' \t', s1.strip().lower(), s2.strip().lower())
    return s.ratio()

def strings_equal(s1, s2):
    if type(s1) == str and type(s2) == str:
        return compare_strings(s1, s2) > 0.8
    return False

def match_strings(string, strings_to_match):
    string_ratios = [{'string': s2, 'ratio': compare_strings(string, s2)} for s2 in strings_to_match]
    best_match = max(string_ratios, key=lambda x: x['ratio'])
    if strings_equal(string, best_match['string']):
        return best_match['string']
    return None