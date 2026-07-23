use super::helpers::run_lua_one;

#[test]
fn test_control_if_nested_priority_top_true_nested_true() {
    assert_eq!(
        run_lua_one(
            "local x=10; local y=''; if x > 0 then if x < 20 then y='inner' else y='outer' end else y='no' end; print(y)"
        ),
        "inner"
    );
}

#[test]
fn test_control_if_nested_priority_top_true_nested_false() {
    assert_eq!(
        run_lua_one(
            "local x=30; local y=''; if x > 0 then if x < 20 then y='inner' else y='outer' end else y='no' end; print(y)"
        ),
        "outer"
    );
}

#[test]
fn test_control_if_nested_priority_top_false() {
    assert_eq!(
        run_lua_one(
            "local x=-1; local y=''; if x > 0 then if x < 20 then y='inner' else y='outer' end else y='no' end; print(y)"
        ),
        "no"
    );
}

#[test]
fn test_control_if_nested_priority_nested_chain() {
    assert_eq!(
        run_lua_one(
            "local x=1; local y=''; if x == 0 then y='z' elseif x == 1 then if x == 1 then y='one' else y='other' end else y='x' end; print(y)"
        ),
        "one"
    );
}

#[test]
fn test_control_if_nested_priority_multiple_levels() {
    assert_eq!(
        run_lua_one(
            "local x=2; local y=''; if x > 0 then if x > 1 then if x > 2 then y='high' else y='mid' end else y='low' end else y='neg' end; print(y)"
        ),
        "mid"
    );
}

#[test]
fn test_control_if_nested_priority_three_levels_false() {
    assert_eq!(
        run_lua_one(
            "local x=5; local y=''; if x > 0 then if x > 1 then if x > 10 then y='high' else y='mid' end else y='low' end else y='neg' end; print(y)"
        ),
        "mid"
    );
}

#[test]
fn test_control_if_nested_priority_three_levels_outer() {
    assert_eq!(
        run_lua_one(
            "local x=-5; local y=''; if x > 0 then if x > 1 then if x > 10 then y='high' else y='mid' end else y='low' end else y='neg' end; print(y)"
        ),
        "neg"
    );
}

#[test]
fn test_control_if_nested_priority_no_nested_when_outer_false() {
    assert_eq!(
        run_lua_one(
            "local x=0; local y=''; if x > 0 then if x < 1 then y='inner' end else y='outer' end; print(y)"
        ),
        "outer"
    );
}

#[test]
fn test_control_if_nested_priority_inner_else() {
    assert_eq!(
        run_lua_one(
            "local x=1; local y=''; if x > 0 then if x < 0 then y='inner' else y='inner_else' end else y='outer' end; print(y)"
        ),
        "inner_else"
    );
}

#[test]
fn test_control_if_nested_priority_parentheses_blocking() {
    assert_eq!(
        run_lua_one(
            "local x=2; local y=''; if (x > 0) and (x < 10) then y='ok' else y='bad' end; print(y)"
        ),
        "ok"
    );
}

#[test]
fn test_control_if_nested_priority_priority_and_true() {
    assert_eq!(
        run_lua_one(
            "local x=2; local y=''; if x > 0 and x < 10 and x == 2 then y='ok' else y='bad' end; print(y)"
        ),
        "ok"
    );
}

#[test]
fn test_control_if_nested_priority_priority_or() {
    assert_eq!(
        run_lua_one(
            "local x=12; local y=''; if x < 2 or x > 10 then y='ok' else y='bad' end; print(y)"
        ),
        "ok"
    );
}

#[test]
fn test_control_if_nested_priority_not_true() {
    assert_eq!(
        run_lua_one("local x=0; local y=''; if not (x > 0) then y='ok' else y='bad' end; print(y)"),
        "ok"
    );
}

#[test]
fn test_control_if_nested_priority_not_false() {
    assert_eq!(
        run_lua_one("local x=1; local y=''; if not (x > 0) then y='bad' else y='ok' end; print(y)"),
        "ok"
    );
}

#[test]
fn test_control_if_nested_priority_nested_elseif_chain() {
    assert_eq!(
        run_lua_one(
            "local x=2; local y=''; if x < 0 then y='neg' elseif x == 0 then y='zero' elseif x == 1 then y='one' else y='other' end; print(y)"
        ),
        "other"
    );
}

#[test]
fn test_control_if_nested_priority_nested_elseif_zero() {
    assert_eq!(
        run_lua_one(
            "local x=0; local y=''; if x < 0 then y='neg' elseif x == 0 then y='zero' else y='other' end; print(y)"
        ),
        "zero"
    );
}

#[test]
fn test_control_if_nested_priority_nested_elseif_one() {
    assert_eq!(
        run_lua_one(
            "local x=1; local y=''; if x < 0 then y='neg' elseif x == 0 then y='zero' elseif x == 1 then y='one' else y='other' end; print(y)"
        ),
        "one"
    );
}

#[test]
fn test_control_if_nested_priority_local_scope_shadow() {
    assert_eq!(
        run_lua_one("local x=1; if true then local x='inner'; print(x) end; print(x)"),
        "1"
    );
}

#[test]
fn test_control_if_nested_priority_outer_scope_ref() {
    assert_eq!(
        run_lua_one("local x=1; if true then local y=x; x='inner'; print(y) end; print(x)"),
        "1"
    );
}

#[test]
fn test_control_if_nested_priority_true_false_chain() {
    assert_eq!(
        run_lua_one(
            "local x=0; local y=''; if false then y='no' elseif true then y='yes' else y='never' end; print(y)"
        ),
        "yes"
    );
}
