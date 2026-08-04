-- vybe-test: lua/metatables_newindex/test_newindex_proxy_table
-- origin: languages/lua/tests/lua/test_metatables_newindex.rs

local __w1 = "1 nil"
local __i = 0

local data={}; local proxy=setmetatable({}, {__index=data, __newindex=function(t,k,v) if type(v)=='number' then data[k]=v end end}); proxy.a=1; proxy.b='str'; do local __t = tostring(proxy.a..' '..(proxy.b or 'nil')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
