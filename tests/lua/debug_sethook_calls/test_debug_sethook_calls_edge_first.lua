-- vybe-test: lua/debug_sethook_calls/test_debug_sethook_calls_edge_first
-- origin: languages/lua/tests/lua/test_debug_sethook_calls.rs

local __w1 = "true"
local __i = 0

local n = 0
debug.sethook(function()
  n = n + 1
end, "c")
local function f()
  local s = 0
  for i = 1, 16 do s = s + i end
  return s
end
f()
debug.sethook()
do local __t = tostring(n >= 1); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
