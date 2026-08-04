-- vybe-test: lua/metatables_proxy_tables/test_proxy_tracking_access
-- origin: languages/lua/tests/lua/test_metatables_proxy_tables.rs

local __w1 = "30 30 3"
local __i = 0

local function Track(t)
    local proxy = {}
    local accesses = 0
    local mt = {
        __index = function(_, k)
            accesses = accesses + 1
            return t[k]
        end,
        __newindex = function(_, k, v)
            t[k] = v
        end
    }
    setmetatable(proxy, mt)
    return proxy, function() return accesses end
end
local p, get_accesses = Track({x = 10, y = 20})
local a = p.x + p.y
p.z = 30
do local __t = tostring(a .. ' ' .. p.z .. ' ' .. get_accesses()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
