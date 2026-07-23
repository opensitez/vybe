use super::helpers::run_lua_one;

#[test]
fn test_assert_message_alpha() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "alpha") end); print(ok == false and string.find(tostring(err), "alpha") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_assert_message_beta() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "beta") end); print(ok == false and string.find(tostring(err), "beta") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_assert_message_gamma() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "gamma") end); print(ok == false and string.find(tostring(err), "gamma") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_assert_message_delta() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "delta") end); print(ok == false and string.find(tostring(err), "delta") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_assert_message_epsilon() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "epsilon") end); print(ok == false and string.find(tostring(err), "epsilon") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_assert_message_zeta() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "zeta") end); print(ok == false and string.find(tostring(err), "zeta") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_assert_message_eta() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "eta") end); print(ok == false and string.find(tostring(err), "eta") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_assert_message_theta() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "theta") end); print(ok == false and string.find(tostring(err), "theta") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_assert_message_iota() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "iota") end); print(ok == false and string.find(tostring(err), "iota") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_assert_message_kappa() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "kappa") end); print(ok == false and string.find(tostring(err), "kappa") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_assert_message_lambda() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "lambda") end); print(ok == false and string.find(tostring(err), "lambda") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_assert_message_mu() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "mu") end); print(ok == false and string.find(tostring(err), "mu") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_assert_message_nu() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "nu") end); print(ok == false and string.find(tostring(err), "nu") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_assert_message_xi() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "xi") end); print(ok == false and string.find(tostring(err), "xi") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_assert_message_omicron() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "omicron") end); print(ok == false and string.find(tostring(err), "omicron") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_assert_message_pi() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "pi") end); print(ok == false and string.find(tostring(err), "pi") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_assert_message_rho() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "rho") end); print(ok == false and string.find(tostring(err), "rho") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_assert_message_sigma() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "sigma") end); print(ok == false and string.find(tostring(err), "sigma") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_assert_message_tau() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "tau") end); print(ok == false and string.find(tostring(err), "tau") ~= nil)"#
        ),
        "true"
    );
}

#[test]
fn test_assert_message_upsilon() {
    assert_eq!(
        run_lua_one(
            r#"local ok, err = pcall(function() assert(false, "upsilon") end); print(ok == false and string.find(tostring(err), "upsilon") ~= nil)"#
        ),
        "true"
    );
}
