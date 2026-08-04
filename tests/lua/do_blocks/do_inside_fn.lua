-- vybe-test: lua/do_blocks/do_inside_fn
-- origin: languages/lua/tests/lua/test_do_blocks.rs

local __w1 = "3"
local __i = 0

local function f()
  local x = 1
  do
    local x = 2
    do
      local x = 3
      return x
    end
  end
end
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
