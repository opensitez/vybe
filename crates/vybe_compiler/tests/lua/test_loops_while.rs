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
    }
}
