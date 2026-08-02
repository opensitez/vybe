# vybe-test: python/python_struct_byte_order/test_struct_iter_unpack
# origin: languages/python/tests/python/test_python_struct_byte_order.rs

import struct
data = struct.pack(">HHH", 10, 20, 30)
values = [v for (v,) in struct.iter_unpack(">H", data)]
print(values)
