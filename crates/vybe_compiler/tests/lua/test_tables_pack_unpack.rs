lua_print! {
    test_pack_basic => { "local t = table.pack(1, 2, 3); print(t.n..' '..t[1]..' '..t[2]..' '..t[3])", "3 1 2 3" },
    test_pack_empty => { "local t = table.pack(); print(t.n)", "0" },
    test_pack_nil => { "local t = table.pack(1, nil, 3); print(t.n..' '..t[1]..' '..tostring(t[2])..' '..t[3])", "3 1 nil 3" },
    test_unpack_basic => { "local a, b, c = table.unpack({1, 2, 3}); print(a..' '..b..' '..c)", "1 2 3" },
    test_unpack_range => { "local a, b = table.unpack({1, 2, 3, 4}, 2, 3); print(a..' '..b)", "2 3" },
    test_unpack_empty => { "local a = table.unpack({}); print(tostring(a))", "nil" },
    test_unpack_invalid_range => { "local ok, err = pcall(function() table.unpack({}, 2, 1) end); print(tostring(ok))", "true" },
    test_unpack_large_range => { "local ok, err = pcall(function() table.unpack({}, 1, 100000000) end); print(tostring(ok))", "false" }
}
