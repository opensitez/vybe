# vybe-test: python/py_control_flow_generators_iterators/test_py_generator_recursion_tree_walk
# origin: languages/python/tests/python/test_py_control_flow_generators_iterators.rs

tree = [1, [2, [3, 4]], 5]

def flatten(nested):
    for item in nested:
        if isinstance(item, list):
            yield from flatten(item)
        else:
            yield item

print(list(flatten(tree)))
