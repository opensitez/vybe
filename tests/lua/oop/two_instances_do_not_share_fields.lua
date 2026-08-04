-- vybe-test: lua/oop/two_instances_do_not_share_fields
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "0"
local __i = 0

local A = {}
function A.new() return {n = 0} end
local a = A.new()
local b = A.new()
a.n = 3
do local __t = tostring(b.n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
