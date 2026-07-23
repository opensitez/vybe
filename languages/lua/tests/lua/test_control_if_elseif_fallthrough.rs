use super::helpers::run_lua_one;

#[test]
fn test_control_if_elseif_fallthrough_top() {
    assert_eq!(
        run_lua_one(
            "local x=1; local y=''; if x == 1 then y='one' elseif x == 2 then y='two' else y='other' end; print(y)"
        ),
        "one"
    );
}

#[test]
fn test_control_if_elseif_fallthrough_second() {
    assert_eq!(
        run_lua_one(
            "local x=2; local y=''; if x == 1 then y='one' elseif x == 2 then y='two' else y='other' end; print(y)"
        ),
        "two"
    );
}

#[test]
fn test_control_if_elseif_fallthrough_default() {
    assert_eq!(
        run_lua_one(
            "local x=3; local y=''; if x == 1 then y='one' elseif x == 2 then y='two' else y='other' end; print(y)"
        ),
        "other"
    );
}

#[test]
fn test_control_if_elseif_fallthrough_no_else() {
    assert_eq!(
        run_lua_one(
            "local x=1; local y=''; if x == 1 then y='one' elseif x == 2 then y='two' end; print(y)"
        ),
        "one"
    );
}

#[test]
fn test_control_if_elseif_fallthrough_last_elseif() {
    assert_eq!(
        run_lua_one(
            "local x=3; local y=''; if x == 1 then y='one' elseif x == 2 then y='two' elseif x == 3 then y='three' else y='other' end; print(y)"
        ),
        "three"
    );
}

#[test]
fn test_control_if_elseif_fallthrough_multiple_true_like() {
    assert_eq!(
        run_lua_one(
            "local x=1; local y=''; if x >= 1 then y='big' elseif x > 1 then y='bigger' else y='none' end; print(y)"
        ),
        "big"
    );
}

#[test]
fn test_control_if_elseif_fallthrough_boolean_chain() {
    assert_eq!(
        run_lua_one(
            "local x=0; local y=''; if x > 0 then y='a' elseif x < 0 then y='b' else y='c' end; print(y)"
        ),
        "c"
    );
}

#[test]
fn test_control_if_elseif_fallthrough_nested_elseif() {
    assert_eq!(
        run_lua_one(
            "local x=4; local y=''; if x < 0 then y='neg' elseif x == 1 then y='one' elseif x == 2 then y='two' elseif x == 3 then y='three' else y='other' end; print(y)"
        ),
        "other"
    );
}

#[test]
fn test_control_if_elseif_fallthrough_nested_elseif_match() {
    assert_eq!(
        run_lua_one(
            "local x=3; local y=''; if x < 0 then y='neg' elseif x == 1 then y='one' elseif x == 2 then y='two' elseif x == 3 then y='three' else y='other' end; print(y)"
        ),
        "three"
    );
}

#[test]
fn test_control_if_elseif_fallthrough_middle_default() {
    assert_eq!(
        run_lua_one(
            "local x=10; local y=''; if x < 0 then y='neg' elseif x == 1 then y='one' elseif x == 2 then y='two' elseif x == 3 then y='three' else y='other' end; print(y)"
        ),
        "other"
    );
}

#[test]
fn test_control_if_elseif_fallthrough_expression() {
    assert_eq!(
        run_lua_one(
            "local x=4; local y=''; if (x * 0) == 0 and x == 4 then y='ok' elseif x > 0 then y='pos' else y='no' end; print(y)"
        ),
        "ok"
    );
}

#[test]
fn test_control_if_elseif_fallthrough_not_expression() {
    assert_eq!(
        run_lua_one(
            "local x=4; local y=''; if not (x < 0) then y='ok' elseif x == 4 then y='four' else y='no' end; print(y)"
        ),
        "ok"
    );
}

#[test]
fn test_control_if_elseif_fallthrough_and_expression() {
    assert_eq!(
        run_lua_one(
            "local x=4; local y=''; if x > 0 and x < 5 then y='ok' elseif x > 5 then y='big' else y='no' end; print(y)"
        ),
        "ok"
    );
}

#[test]
fn test_control_if_elseif_fallthrough_or_expression() {
    assert_eq!(
        run_lua_one(
            "local x=4; local y=''; if x > 10 or x == 4 then y='ok' elseif x == 2 then y='two' else y='no' end; print(y)"
        ),
        "ok"
    );
}

#[test]
fn test_control_if_elseif_fallthrough_nested_blocks() {
    assert_eq!(
        run_lua_one(
            "local x=2; local y=''; if x > 1 then if x == 1 then y='one' else y='not_one' end else y='low' end; print(y)"
        ),
        "not_one"
    );
}

#[test]
fn test_control_if_elseif_fallthrough_default_shadow() {
    assert_eq!(
        run_lua_one(
            "local x=''; local y=''; if x == nil then y='nil' elseif x == '' then y='empty' else y='other' end; print(y)"
        ),
        "nil"
    );
}

#[test]
fn test_control_if_elseif_fallthrough_nil_check() {
    assert_eq!(
        run_lua_one(
            "local x=nil; local y=''; if x == 1 then y='one' elseif x == nil then y='nil' else y='other' end; print(y)"
        ),
        "nil"
    );
}

#[test]
fn test_control_if_elseif_fallthrough_string_case() {
    assert_eq!(
        run_lua_one(
            "local x='a'; local y=''; if x == 'x' then y='x' elseif x == 'a' then y='a' elseif x == 'b' then y='b' else y='other' end; print(y)"
        ),
        "a"
    );
}

#[test]
fn test_control_if_elseif_fallthrough_number_chain() {
    assert_eq!(
        run_lua_one(
            "local x=9; local y=''; if x == 1 then y='one' elseif x == 3 then y='three' elseif x == 9 then y='nine' else y='other' end; print(y)"
        ),
        "nine"
    );
}

#[test]
fn test_control_if_elseif_fallthrough_last_else() {
    assert_eq!(
        run_lua_one(
            "local x=42; local y=''; if x > 100 then y='high' elseif x > 10 then y='mid' else y='low' end; print(y)"
        ),
        "mid"
    );
}
