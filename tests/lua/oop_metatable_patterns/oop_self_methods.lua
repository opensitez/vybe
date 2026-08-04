-- vybe-test: lua/oop_metatable_patterns/oop_self_methods
-- origin: languages/lua/tests/lua/test_oop_metatable_patterns.rs

local __w1 = "100"
local __i = 0

local Account = {balance = 0}
Account.__index = Account
function Account:deposit(v) self.balance = self.balance + v end
local a = setmetatable({}, Account)
a:deposit(100)
do local __t = tostring(a.balance); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
