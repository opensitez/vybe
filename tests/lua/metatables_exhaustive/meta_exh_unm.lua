-- vybe-test: lua/metatables_exhaustive/meta_exh_unm
-- origin: languages/lua/tests/lua/test_metatables_exhaustive.rs

local __w1 = "-5"
local __i = 0

local mt = {__unm = function(a) return -a.v end}
local function W(v) return setmetatable({v=v}, mt) end
do local __t = tostring(-W(5)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
