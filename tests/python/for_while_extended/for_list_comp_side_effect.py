# vybe-test: python/for_while_extended/for_list_comp_side_effect
# origin: languages/python/tests/python/test_for_while_extended.rs

out = []
for x in range(3):
 out.append(x)
print(out)
