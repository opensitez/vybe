-- vybe-test: lua/do_blocks/do_block_nested_local_function_shadows
-- origin: languages/lua/tests/lua/test_do_blocks.rs

local __w1 = "inner\nouter"
local __i = 0

local function f() return "outer" end
do
  local function f() return "inner" end
  do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
