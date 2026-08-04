-- vybe-test: lua/oop/private_state_via_closure_in_constructor
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "12"
local __i = 0

local function make_counter(init)
  local n = init
  return {
    get = function() return n end,
    inc = function() n = n + 1 end,
  }
end
local c = make_counter(10)
c.inc(); c.inc()
do local __t = tostring(c.get()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
