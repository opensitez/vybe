lua_print! {
    test_repeat_basic => {
        "local i=1; local s=''; repeat s=s..i; i=i+1 until i>3; print(s)",
        "123"
    },
    test_repeat_executes_at_least_once => {
        "local i=10; local s=''; repeat s=s..i until true; print(s)",
        "10"
    },
    test_repeat_scope_of_until => {
        "local i=1; local s=''; repeat local j=i; s=s..j; i=i+1 until j>=3; print(s)",
        "123"
    },
    test_repeat_break => {
        "local i=1; local s=''; repeat s=s..i; if i==2 then break end; i=i+1 until false; print(s)",
        "12"
    },
    test_repeat_nested => {
        "local i=1; local s=''; repeat local j=1; repeat s=s..i..j; j=j+1 until j>2; i=i+1 until i>2; print(s)",
        "11122122"
    },
    test_repeat_closure_in_until => {
        "local i=1; repeat local j=i; local f=function() return j>2 end; i=i+1 until f(); print(i)",
        "4"
    },
    test_repeat_truthiness => {
        "local s=''; local t={1, 2, nil}; local i=1; repeat s=s..t[i]; i=i+1 until not t[i]; print(s)",
        "12"
    }
}
