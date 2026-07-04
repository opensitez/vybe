lua_print! {
    test_mixed_table_length_operator_no_holes => { "local t={1, 2, 3, a=10, b=20}; print(#t)", "3" },
    test_mixed_table_length_operator_hole_at_end => { "local t={1, 2, nil, a=10}; print(#t)", "2" },
    test_mixed_table_length_operator_hole_in_middle => { "local t={1, nil, 3, a=10}; local len=#t; print(len==1 or len==3)", "true" },
    test_mixed_table_iteration_pairs => { "local t={1, a=2}; local c=0; for k,v in pairs(t) do c=c+1 end; print(c)", "2" },
    test_mixed_table_iteration_ipairs => { "local t={1, 2, a=3}; local c=0; for k,v in ipairs(t) do c=c+1 end; print(c)", "2" },
    test_mixed_table_ipairs_stops_at_hole => { "local t={1, nil, 3, a=4}; local c=0; for k,v in ipairs(t) do c=c+1 end; print(c)", "1" },
    test_mixed_table_next_includes_all => { "local t={1, a=2}; local c=0; local k=nil; while true do k = next(t, k); if not k then break end; c=c+1 end; print(c)", "2" },
    test_mixed_table_insert => { "local t={a=1}; table.insert(t, 10); print(t[1]..' '..t.a)", "10 1" },
    test_mixed_table_remove => { "local t={10, a=1}; local v=table.remove(t); print(v..' '..(t[1] or 'nil')..' '..t.a)", "10 nil 1" },
    test_mixed_table_move => { "local t={10, 20, a=1}; local t2={b=2}; table.move(t, 1, 2, 1, t2); print(t2[1]..' '..t2[2]..' '..(t2.a or 'nil')..' '..t2.b)", "10 20 nil 2" },
    test_mixed_table_unpack => { "local t={10, 20, a=1}; local a, b = table.unpack(t); print(a..' '..b)", "10 20" },
    test_mixed_table_concat => { "local t={10, 20, a=1}; print(table.concat(t, ','))", "10,20" },
    test_mixed_table_pack => { "local t=table.pack(1, 2); t.a=3; print(t.n..' '..t[1]..' '..t.a)", "2 1 3" },
    test_mixed_table_constructor_order => { "local t={a=1, 10, b=2, 20}; print(t[1]..' '..t[2]..' '..t.a..' '..t.b)", "10 20 1 2" },
    test_mixed_table_constructor_overwrite => { "local t={10, [1]=20}; print(t[1])", "20" },
    test_mixed_table_constructor_overwrite_reverse => { "local t={[1]=20, 10}; print(t[1])", "10" }
}
