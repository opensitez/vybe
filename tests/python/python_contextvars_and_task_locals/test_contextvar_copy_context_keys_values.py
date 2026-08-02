# vybe-test: python/python_contextvars_and_task_locals/test_contextvar_copy_context_keys_values
# origin: languages/python/tests/python/test_python_contextvars_and_task_locals.rs

import contextvars

v1 = contextvars.ContextVar('v1')
v2 = contextvars.ContextVar('v2')

v1.set(100)
v2.set('hello')

ctx = contextvars.copy_context()
items = {k.name: v for k, v in ctx.items()}
print(items['v1'], items['v2'])
