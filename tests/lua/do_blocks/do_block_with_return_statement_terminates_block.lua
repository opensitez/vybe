-- vybe-test: lua/do_blocks/do_block_with_return_statement_terminates_block
-- origin: languages/lua/tests/lua/test_do_blocks.rs

local __w1 = "early"
local __i = 0

local function f()
  do
    return "early"
  end
  return "late"
end
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
