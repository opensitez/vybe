-- vybe-test: lua/metatables/__tostring_metamethod_used_by_tostring
-- origin: languages/lua/tests/lua/test_metatables.rs

local __w1 = "tbl"
local __i = 0

local t=setmetatable({},{__tostring=function() return "tbl" end})
do local __t = tostring(tostring(t)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
