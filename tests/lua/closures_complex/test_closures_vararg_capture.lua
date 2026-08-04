-- vybe-test: lua/closures_complex/test_closures_vararg_capture
-- origin: languages/lua/tests/lua/test_closures_complex.rs

local __w1 = "b"
local __i = 0

local function capture_varargs(...)
    local args = {...}
    return function(i) return args[i] end
end
local f = capture_varargs('a', 'b', 'c')
do local __t = tostring(f(2)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
