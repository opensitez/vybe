-- vybe-test: lua/modules_complex/test_module_cyclic_dependency
-- origin: languages/lua/tests/lua/test_modules_complex.rs

local __w1 = "true true"
local __i = 0

local m1, m2 = {}, {}
        package.loaded['m1'] = m1
        package.loaded['m2'] = m2
        m1.get_m2 = function() return require('m2') end
        m2.get_m1 = function() return require('m1') end
        do local __t = tostring(tostring(m1.get_m2() == m2) .. ' ' .. tostring(m2.get_m1() == m1)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
