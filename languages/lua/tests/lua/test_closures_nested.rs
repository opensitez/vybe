lua_print! {
    test_nested_closure_deep_read => { "local function f1(a) return function(b) return function(c) return function(d) return a..b..c..d end end end end; print(f1(1)(2)(3)(4))", "1234" },
    test_nested_closure_deep_write => { "local function f1() local a=0; return function() return function() return function() a=a+1; return a end end end end; local f4 = f1()()(); print(f4()..f4())", "12" },
    test_nested_closure_deep_mixed => { "local function f1() local a=10; return function() local b=20; return function() a=a+1; return a+b end end end; local f3 = f1()(); print(f3()..f3())", "3132" },
    test_nested_closure_sibling_functions => { "local function outer() local a=1; local function inner1() a=a+10; return a end; local function inner2() a=a*2; return a end; return function() return inner1()..inner2() end end; print(outer()())", "1122" },
    test_nested_closure_dynamic_depth => { "local function make_adder(n) local sum=n; local function adder(x) if x then sum=sum+x; return adder else return sum end end; return adder end; print(make_adder(1)(2)(3)(4)())", "10" }
}
