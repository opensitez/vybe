pub fn cases() -> Vec<String> {
    let mut cases = [
        "42",
        "(a)+(b)+(c)",
        "(a)..(b)..(c)",
        "(a+b)+(c+d)+(e+f)",
        "(a..b)..(c..d)..(e..f)",
        "a + -- middle\n b + c",
        "a.. -- middle\n b..c",
        "function() function f() end end",
        "function() function t:m() end end",
        "function() function t:m(self,x) return self,x end end",
        "function() function a.b:c(x,...) return self,x,... end end",
        "function() function a:b.c(x) return self,x end end",
        "function() function a:b:c(x) function t:m() end return self,x end end",
        "function() function a -- table\n : b(x) return self,x end end",
        "function() function t:m(x",
        "function() function a.b.c(x,...) return x,... end end",
        "function() function a.b() end function a.c(self) return self end end",
        "function() function a -- base\n . b -- field\n . c(x) return x end end",
        "function() function f(a,...) local x=a return x,... end return f(1) end",
        "function() function f(self) function g(x) return x end return g(self) end end",
        "function() function f(a,...)",
        "0xFF",
        ".5",
        "2e3",
        "1.",
        "nil",
        "true",
        "false",
        "name",
        "...",
        "'café'",
        r#""a\n\x41\u{1f642}\255""#,
        r#""a\z  \n b""#,
        "[=[héllo]=]",
        "[[hello]]",
        "[[\n  hello]]",
        "[[   ]]",
        "[==[\r\n héllo]==]",
        "-2 ^ 2",
        "2 ^ 3 ^ 4",
        "1 + 2 * 3 - 4",
        "not a and b or c",
        "~a ~ b",
        "#t",
        "a << 2 >> 1 & b ~ c | d",
        "a <= b ~= c",
        "a..b..c",
        "a//b % c / d",
        "(1 +\n 2) * 3",
        "f()",
        "f(1, g(2))",
        "print('café')",
        "tostring(1)",
        "error('x')",
        "a.b[2](3).c",
        "t:m(1, f())",
        "t:m().x:y(2)",
        "(f())",
        "((f()))",
        "(f())()",
        "f((g()), (h())())",
        "f('é').m(\r\n 2)",
        "f(1)+",
        "a +",
        "",
        "(",
        "not",
        "/",
    ]
    .into_iter()
    .map(str::to_owned)
    .collect::<Vec<_>>();
    let operators = [
        "+", "-", "*", "/", "//", "%", "^", "..", "<<", ">>", "&", "~", "|", "<", ">", "<=", ">=",
        "~=", "==", "and", "or",
    ];
    cases.extend(
        [
            "{}",
            "{1}",
            "{1,2;3,}",
            "{a=1,b=f(2)}",
            "{a}",
            "{a.b}",
            "{a+1}",
            "{[a]=1,['é']=f()}",
            "{...}",
            "{... + 1}",
            "{... t}",
            "{...t, a, [f()]=g(), x=(h())}",
            "{name -- comment\n = 2, a}",
            "{[[\n text]]}",
            "{[=[key]=]}",
            "{{1}, {a={b=2}}}",
            "f({a=1}, {}).x",
            "{a=1}[f()].x",
            "({1}).x:m({2})",
            "{1} +",
            "{",
            "{a=",
            "{[a]",
            "{... t,",
            "{a,a=}",
        ]
        .into_iter()
        .map(str::to_owned),
    );
    cases.push(format!(
        "{{{}}}",
        (0..200)
            .map(|value| format!("key{value}={value}"))
            .collect::<Vec<_>>()
            .join(",")
    ));
    cases.push(format!("{}1{}", "{".repeat(30), "}".repeat(30)));
    cases.extend(
        [
            "function() end",
            "function(a,b) return a+b end",
            "function(...) return ... end",
            "function(a,...) return a, ... end",
            "function() return end",
            "function() f(1); return 2 end",
            "function() local a return a end",
            "function() local a,b=1,2 return a,b end",
            "function() local a,b=1 return b end",
            "function() local a=1, f(2) return a end",
            "function() local f=function(a) return a end return f(2) end",
            "function() local function f(a,...) return a,... end return f(2) end",
            "{fn=function(a) return {a} end}[1](2)",
            "(function(a) return a end)(1)",
            "function() break end",
            "function() ::label:: goto label end",
            "function() local a,b,c = 1, f(), g() return a,b,c end",
            "function() local a = 1, f(), g() return a end",
            "function(_lua_varargs) return _lua_varargs end",
            "function(a,...) local function f(...) return ... end return f(a,...) end",
            "function() a=1 return a end",
            "function() a,b,c=1,2,3 return a,b,c end",
            "function() a,b=f(),... return a end",
            "function() t.x,t[1]=2,3 end",
            "function() a={1,2}; b,c={3,4},5 end",
            "function() while a do local x=1 a=x break end end",
            "function() while false do end end",
            "function() repeat local x=f() a=x until g(a) end",
            "function() repeat until true end",
            "function() do local x=1 f(x) end end",
            "function() do end end",
            "function() for x in iter do end end",
            "function() for k,v in pairs(t) do f(k,v) end end",
            "function() for k,v,x in iter,state,control do local a=v f(k,x) end end",
            "function() for a,b__lua_extra__c,d in iter do return a,b__lua_extra__c,d end end",
            "function() for x in (f()) do if x then break else g() end end end",
            "function() for k,v in pairs({a=1}) do for x in v do f(k,x) end end end",
            "function() for k,v in pairs(t) do",
            "function() for k,v in do end end",
            "function() if a then f() end end",
            "function() if false then else return 2 end end",
            "function() if a then elseif b then return 2 else return 3 end end",
            "function() if a then f() elseif b then elseif c then g() else end end",
            "function() if a then if b then f() else g() end else h() end end",
            "function() if a then else end end",
            "function() if a then elseif b then end end",
            "function() while a do repeat do a,b=b,a end until f() end end",
            "function() local a -- comment\n = 'é' return a end",
            "function( end",
            "function() local a= end",
            "function() return 1",
        ]
        .into_iter()
        .map(str::to_owned),
    );
    for first in [
        "a", "a=1", "[a]=b", "...t", "...", "... + 1", "{}", "f()", "(f())",
    ] {
        for second in ["a", "b=2", "[b]=a", "...u", "{}", "g()"] {
            cases.push(format!("{{{first}, {second}}}"));
        }
    }
    for first in operators {
        for second in operators {
            cases.push(format!("a {first} b {second} c"));
        }
    }
    let alphabet = ['a', '1', '+', '~', '(', ')'];
    let mut level = vec![String::new()];
    for _ in 0..4 {
        let mut next = Vec::new();
        for prefix in level {
            for suffix in alphabet {
                next.push(format!("{prefix}{suffix}"));
            }
        }
        cases.extend(next.iter().cloned());
        level = next;
    }
    cases
}
