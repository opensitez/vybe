lua_print! {
    test_pack_unpack_i1 => { "local s = string.pack('b', -128); local v, p = string.unpack('b', s); print(v..' '..p)", "-128 2" },
    test_pack_unpack_I1 => { "local s = string.pack('B', 255); local v, p = string.unpack('B', s); print(v..' '..p)", "255 2" },
    test_pack_unpack_i2 => { "local s = string.pack('h', -32768); local v, p = string.unpack('h', s); print(v..' '..p)", "-32768 3" },
    test_pack_unpack_I2 => { "local s = string.pack('H', 65535); local v, p = string.unpack('H', s); print(v..' '..p)", "65535 3" },
    test_pack_unpack_i4 => { "local s = string.pack('i4', -2147483648); local v, p = string.unpack('i4', s); print(v..' '..p)", "-2147483648 5" },
    test_pack_unpack_I4 => { "local s = string.pack('I4', 4294967295); local v, p = string.unpack('I4', s); print(v..' '..p)", "4294967295 5" },
    test_pack_unpack_i8 => { "local s = string.pack('i8', -1); local v, p = string.unpack('i8', s); print(v..' '..p)", "-1 9" },
    test_pack_unpack_f => { "local s = string.pack('f', 3.14); local v, p = string.unpack('f', s); print(tostring(math.abs(v - 3.14) < 0.01))", "true" },
    test_pack_unpack_d => { "local s = string.pack('d', 3.14159265); local v, p = string.unpack('d', s); print(tostring(math.abs(v - 3.14159265) < 0.0001))", "true" },
    test_pack_unpack_string => { "local s = string.pack('z', 'hello'); local v, p = string.unpack('z', s); print(v..' '..p)", "hello 7" },
    test_pack_unpack_fixed_string => { "local s = string.pack('c5', 'hello'); local v, p = string.unpack('c5', s); print(v..' '..p)", "hello 6" },
    test_pack_unpack_sized_string => { "local s = string.pack('s1', 'hello'); local v, p = string.unpack('s1', s); print(v..' '..p)", "hello 7" },
    test_pack_unpack_multiple => { "local s = string.pack('b H z', 42, 1000, 'abc'); local v1, v2, v3, p = string.unpack('b H z', s); print(v1..' '..v2..' '..v3..' '..p)", "42 1000 abc 8" },
    test_pack_little_endian => { "local s = string.pack('< H', 0x1234); local b1, b2 = string.unpack('B B', s); print(b1..' '..b2)", "52 18" },
    test_pack_big_endian => { "local s = string.pack('> H', 0x1234); local b1, b2 = string.unpack('B B', s); print(b1..' '..b2)", "18 52" },
    test_pack_invalid_format => { "local ok = pcall(function() string.pack('q', 1) end); print(tostring(ok))", "false" },
    test_unpack_out_of_bounds => { "local ok = pcall(function() string.unpack('i4', 'ab') end); print(tostring(ok))", "false" }
}
