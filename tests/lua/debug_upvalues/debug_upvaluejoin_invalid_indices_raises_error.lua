-- vybe-test: lua/debug_upvalues/debug_upvaluejoin_invalid_indices_raises_error
-- origin: languages/lua/tests/lua/test_debug_upvalues.rs

local __w1 = "true"
local __i = 0

local function f1() end
local function f2() end
local ok, err = pcall(function() debug.upvaluejoin(f1, 1, f2, 1) end)
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
