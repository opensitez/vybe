-- vybe-test: lua/tables_metatables_oop/test_oop_privacy_closure
-- origin: languages/lua/tests/lua/test_tables_metatables_oop.rs

local __w1 = "80"
local __i = 0

local function make_account(initial)
    local balance = initial
    return {
        withdraw = function(v)
            balance = balance - v
            return balance
        end,
        deposit = function(v)
            balance = balance + v
            return balance
        end,
        get_balance = function()
            return balance
        end
    }
end
local acc = make_account(100)
acc.withdraw(20)
do local __t = tostring(acc.get_balance()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
