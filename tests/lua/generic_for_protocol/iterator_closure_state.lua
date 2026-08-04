-- vybe-test: lua/generic_for_protocol/iterator_closure_state
-- origin: languages/lua/tests/lua/test_generic_for_protocol.rs

local __w1 = "abc"
local __i = 0

local function values(t)
  local i = 0
  return function()
    i = i + 1
    return t[i]
  end
end
local r = ""
for v in values({"a", "b", "c"}) do r = r .. v end
do local __t = tostring(r); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
