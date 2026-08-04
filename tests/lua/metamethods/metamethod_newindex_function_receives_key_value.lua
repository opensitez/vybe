-- vybe-test: lua/metamethods/metamethod_newindex_function_receives_key_value
-- origin: languages/lua/tests/lua/test_metamethods.rs

local __w1 = "x=1"
local __i = 0

local log=""
local t=setmetatable({}, {__newindex=function(_,k,v) log=k.."="..v end})
t.x=1
do local __t = tostring(log); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
