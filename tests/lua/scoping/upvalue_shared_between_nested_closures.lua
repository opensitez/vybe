-- vybe-test: lua/scoping/upvalue_shared_between_nested_closures
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "1"
local __i = 0

local n = 0
local function inc() n = n + 1 end
local function read() return n end
inc()
do local __t = tostring(read()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
