-- vybe-test: lua/programs/sign_function_with_if
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "-1"
local __i = 0

local function sign(n)
  if n > 0 then return 1 elseif n < 0 then return -1 else return 0 end
end
do local __t = tostring(sign(-2)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
