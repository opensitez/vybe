lua_print! {
    test_goto_forward => { "local a=1; goto lbl; a=2; ::lbl::; print(a)", "1" },
    test_goto_backward => { "local a=1; ::lbl::; a=a+1; if a<3 then goto lbl end; print(a)", "3" },
    test_goto_skip_local => { "local a=1; goto lbl; local b=2; ::lbl::; print(a)", "1" },
    test_goto_out_of_block => { "local a=1; do goto lbl end; a=2; ::lbl::; print(a)", "1" },
    test_goto_into_block_error => { "local ok = pcall(function() load('goto lbl; do ::lbl:: end') end); print(tostring(ok))", "false" },
    test_goto_into_scope_of_local_error => { "local ok = pcall(function() load('goto lbl; local a=1; ::lbl::') end); print(tostring(ok))", "false" },
    test_goto_duplicate_label_error => { "local ok = pcall(function() load('::lbl::; ::lbl::') end); print(tostring(ok))", "false" },
    test_goto_break_equivalent => { "local a=1; while true do a=a+1; if a==3 then goto break_lbl end end; ::break_lbl::; print(a)", "3" },
    test_goto_label_shadowing => { "local a=1; ::lbl::; do ::lbl::; a=2; goto end_lbl; ::end_lbl:: end; print(a)", "2" }
}
