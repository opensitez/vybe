pub fn cases() -> Vec<String> {
    let mut cases = [
        "",
        "function t:m(x) return self.x+x end; t:m(1)",
        "function a.b:c(self,...) return self,... end",
        "-- only a comment\n",
        "local x = 1 + 2 * 3; return x",
        "local a,b = f(); return a,b",
        "function f(a, ...) return a, ... end; f(2)",
        "local function f(a) return a+1 end; return f(3)",
        "function a.b.c(x,...) return x,... end; return a.b.c(1)",
        "for k,v in pairs(t) do print(k,v) end",
        "for i=1,10 do print(i) end",
        "for i=3,1,-1 do local i=7 print(i) end",
        "for i=1,2,0 do end",
        "for i=1,3.5 do print(i) end",
        "for i=f(1,2.5),g(3),h(4) do end",
        "for i=9007199254740993,9007199254740995,0.5 do end",
        "for i=1,3 do t[i]=function() return i,__lua_capture_i end end",
        "for i=1,3 do local f=function(i) return i end end",
        "for i=1,3 do local f=function() local i=i return i end end",
        "for i=1,3 do local f=function() return function() return i end end end",
        "for i=1,3 do for j=1,2 do t[i]=function() return i,j end end end",
        "for i=1,3 do if i then local f=function() return i end else local f=function() return i end end end",
        "for i=1,3 do for k,v in pairs(t) do t[k]=function() return i end end end",
        "for a,b__lua_extra__c,d in iter,state,control do return a,d end",
        "if a then f() elseif b then g() else h() end",
        "if false then else return 2 end",
        "while a do a=a-1 end; repeat a=a+1 until a>3",
        "local t = {x=1, [2]=3, 4}; return t.x, t[2]",
        "local s = [==[long string]==]; return #s",
        "local f=function(a) return a end; return f(1)",
        "return not a and b or c, -2^2, a .. b .. c, a ~= b",
        "::again:: goto again",
        "do local x=1 end",
        "return 0xff, 1.25e-2, 'a\\n', true, nil",
        "-- comment\nlocal s='é😀'; return s",
        "local s='é😀'\r\nreturn s",
        "for k,v in pairs(t) do function a.b(x) return x end end",
    ]
    .into_iter()
    .map(str::to_owned)
    .collect::<Vec<_>>();
    cases.push(
        (0..200)
            .map(|n| format!("local x{n} = {n} + f(1,2*3)\n"))
            .collect(),
    );
    cases
}
