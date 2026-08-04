-- vybe-test: lua/base_error_wrapping/test_error_chained_calls
-- origin: languages/lua/tests/lua/test_base_error_wrapping.rs

local __w1 = "true"
local __i = 0

local function a() return b() end
local function b() return c() end
local function c() error("chain") end
local ok, err = pcall(a)
do local __t = tostring(string.find(err, "chain") ~= nil); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
