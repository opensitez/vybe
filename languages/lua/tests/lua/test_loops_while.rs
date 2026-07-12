lua_print! {
    test_while_basic => {
        "local i=1; local s=''; while i<=3 do s=s..i; i=i+1 end; print(s)",
        "123"
    },
    test_while_condition_false_initially => {
        "local i=10; while i<5 do i=i+1 end; print(i)",
        "10"
    },
    test_while_break => {
        "local i=1; local s=''; while true do s=s..i; if i==3 then break end; i=i+1 end; print(s)",
        "123"
    },
    test_while_nested => {
        "local i=1; local s=''; while i<=2 do local j=1; while j<=2 do s=s..i..j; j=j+1 end; i=i+1 end; print(s)",
        "11122122"
    },
    test_while_break_nested => {
        "local i=1; local s=''; while i<=3 do local j=1; while j<=3 do if j==2 then break end; s=s..i..j; j=j+1 end; i=i+1 end; print(s)",
        "112131"
    },
    test_while_truthiness => {
        "local i=3; local s=''; while i do s=s..i; i=i-1; if i==0 then i=nil end end; print(s)",
        "321"
    },
    test_while_false_truthiness => {
        "local flag=false; local cnt=0; while flag do cnt=cnt+1 end; print(cnt)",
        "0"
    },
    test_while_multiple_conditions => {
        "local a=1; local b=5; local s=''; while a<3 and b>3 do s=s..a..b; a=a+1; b=b-1 end; print(s)",
        "1524"
    },
    test_while_local_scoping => {
        "local s=''; local x=0; while x<2 do local y=x; x=x+1; s=s..y end; print(s)",
        "01"
    },
    test_while_closure_capture => {
        "local f; local i=1; while i<=2 do local j=i; if i==1 then f = function() return j end end; i=i+1 end; print(f())",
        "1"
    },
    while_with_complex_condition_short_circuit => {
        "local i, count = 1, 0\nwhile i <= 5 and (function() count = count + 1; return i % 2 == 0 end)() do\n  i = i + 1\nend\nprint(i .. \",\" .. count)\n",
        "1,1"
    },
    while_shadowing_outer_control_variable => {
        "local i = 10\nwhile i > 0 do\n  local i = 5\n  print(i)\n  break\nend\n",
        "5"
    },
    while_break_with_returned_values_from_block => {
        "local function f()\n  local i = 1\n  while true do\n    if i == 5 then return i end\n    i = i + 1\n  end\nend\nprint(f())\n",
        "5"
    },
    while_empty_body => {
        "local i = 1\nwhile (function() i = i + 1; return i < 3 end)() do end\nprint(i)\n",
        "3"
    },
    while_condition_mutates_upvalue_shared_with_body => {
        "local x = 10\nlocal function dec() x = x - 1; return x > 5 end\nlocal sum = 0\nwhile dec() do\n  sum = sum + x\nend\nprint(sum)\n",
        "30"
    },
    while_nested_control_with_labels_and_gotos => {
        "local i = 0\nlocal sum = 0\nwhile i < 3 do\n  i = i + 1\n  local j = 0\n  while j < 3 do\n    j = j + 1\n    if j == 2 then goto next_outer end\n    sum = sum + (i * 10 + j)\n  end\n  ::next_outer::\nend\nprint(sum)\n",
        "33"
    },
    while_conditional_break_with_nil_and_false => {
        "local val = false\nlocal sum = 0\nwhile not val do\n  sum = sum + 1\n  if sum == 3 then val = nil end\n  if sum == 5 then break end\nend\nprint(sum)\n",
        "5"
    },
    while_upvalues_accumulated_in_table_within_loop => {
        "local tbl = {}\nlocal i = 1\nwhile i <= 3 do\n  local x = i * 10\n  tbl[i] = function() return x end\n  i = i + 1\nend\nprint(tbl[1]() .. \",\" .. tbl[2]() .. \",\" .. tbl[3]())\n",
        "10,20,30"
    }
}
