-- vybe-test: lua/metatables_proxy_tables/test_proxy_default_value
-- origin: languages/lua/tests/lua/test_metatables_proxy_tables.rs

local __w1 = "1 0"
local __i = 0

local function DefaultTable(default)
    local mt = {
        __index = function(t, k)
            return default
        end
    }
    return setmetatable({}, mt)
end
local counts = DefaultTable(0)
counts['a'] = counts['a'] + 1
do local __t = tostring(counts['a'] .. ' ' .. counts['b']); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
