lua_print! {
    test_math_random_large => { "local r = math.random(1000000000); print(type(r) == 'number' and r >= 1 and r <= 1000000000)", "true" },
    test_math_random_negative_range => { "local r = math.random(-10, -1); print(type(r) == 'number' and r >= -10 and r <= -1)", "true" },
    test_math_random_same_range => { "local r = math.random(5, 5); print(r)", "5" },
    test_math_random_invalid_range_negative => { "local ok = pcall(function() math.random(5, -5) end); print(tostring(ok))", "false" },
    test_math_randomseed_zero => { "math.randomseed(0); local r1 = math.random(); math.randomseed(0); local r2 = math.random(); print(tostring(r1 == r2))", "true" },
    test_math_randomseed_negative => { "math.randomseed(-1); local r1 = math.random(); math.randomseed(-1); local r2 = math.random(); print(tostring(r1 == r2))", "true" }
}
