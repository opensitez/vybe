-- vybe-test: lua/metamethods/metamethod_len_returns_custom
-- origin: languages/lua/tests/lua/test_metamethods.rs

local __w1 = "99"
local __i = 0

local mt={__len=function() return 99 end}
do local __t = tostring(#setmetatable({},mt)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
