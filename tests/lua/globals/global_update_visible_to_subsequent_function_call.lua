-- vybe-test: lua/globals/global_update_visible_to_subsequent_function_call
-- origin: languages/lua/tests/lua/test_globals.rs

local __w1 = "3"
local __i = 0

shared = 0
local function inc() shared = shared + 1 end
local function get() return shared end
inc(); inc(); inc()
do local __t = tostring(get()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
