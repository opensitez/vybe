-- vybe-test: lua/errors/xpcall_passes_extra_args_to_target
-- origin: languages/lua/tests/lua/test_errors.rs

local __w1 = "handled:equal"
local __i = 0

local function f(x, y) if x == y then error("equal") end end
local ok, msg = xpcall(f, function(e) return "handled:"..e end, 5, 5)
do local __t = tostring(msg); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
