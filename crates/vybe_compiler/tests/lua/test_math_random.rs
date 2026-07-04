lua_print! {
    test_math_random_no_args => { "local r = math.random(); print(type(r) == 'number' and r >= 0 and r < 1)", "true" },
    test_math_random_one_arg => { "local r = math.random(5); print(type(r) == 'number' and r >= 1 and r <= 5 and math.floor(r) == r)", "true" },
    test_math_random_two_args => { "local r = math.random(10, 20); print(type(r) == 'number' and r >= 10 and r <= 20 and math.floor(r) == r)", "true" },
    test_math_random_invalid_args => { "local ok = pcall(function() math.random(20, 10) end); print(tostring(ok))", "false" },
    test_math_randomseed => { "math.randomseed(42); local r1 = math.random(); math.randomseed(42); local r2 = math.random(); print(tostring(r1 == r2))", "true" }
}
