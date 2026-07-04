lua_print! {
    test_and_true_true => { "print(tostring(true and true))", "true" },
    test_and_true_false => { "print(tostring(true and false))", "false" },
    test_and_false_true => { "print(tostring(false and true))", "false" },
    test_and_short_circuit => { "local a=0; local _ = false and (function() a=1 return true end)(); print(a)", "0" },
    test_and_return_second => { "print(tostring(10 and 20))", "20" },
    test_and_return_first_false => { "print(tostring(nil and 20))", "nil" },
    test_or_true_false => { "print(tostring(true or false))", "true" },
    test_or_false_false => { "print(tostring(false or false))", "false" },
    test_or_short_circuit => { "local a=0; local _ = true or (function() a=1 return true end)(); print(a)", "0" },
    test_or_return_first => { "print(tostring(10 or 20))", "10" },
    test_or_return_second_false => { "print(tostring(nil or 20))", "20" },
    test_not_true => { "print(tostring(not true))", "false" },
    test_not_false => { "print(tostring(not false))", "true" },
    test_not_nil => { "print(tostring(not nil))", "true" },
    test_not_number => { "print(tostring(not 10))", "false" }
}
