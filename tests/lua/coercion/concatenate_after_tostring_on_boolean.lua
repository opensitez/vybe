-- vybe-test: lua/coercion/concatenate_after_tostring_on_boolean
-- origin: languages/lua/tests/lua/test_coercion.rs

local __w1 = "ok=true"
local __i = 0

do local __t = tostring("ok=" .. tostring(true)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
