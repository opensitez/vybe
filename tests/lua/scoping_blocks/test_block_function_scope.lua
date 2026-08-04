-- vybe-test: lua/scoping_blocks/test_block_function_scope
-- origin: languages/lua/tests/lua/test_scoping_blocks.rs

local __w1 = "42"
local __i = 0

local f; do local a=42; f = function() return a end end; do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
