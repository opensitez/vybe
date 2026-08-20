-- App 23: Ledger and Reconciliation

local accounts = {
  cash = 1000,
  revenue = 0,
  expense = 0,
  receivable = 0,
}

local entries = {}

local function post(desc, debit, credit, amount)
  assert(accounts[debit] ~= nil and accounts[credit] ~= nil, "unknown account")
  assert(amount > 0, "amount must be positive")
  accounts[debit] = accounts[debit] + amount
  accounts[credit] = accounts[credit] - amount
  entries[#entries + 1] = {desc = desc, debit = debit, credit = credit, amount = amount}
end

post("Invoice #100", "receivable", "revenue", 300)
post("Payment #100", "cash", "receivable", 300)
post("Hosting bill", "expense", "cash", 120)

local function trial_balance()
  local total = 0
  for name, bal in pairs(accounts) do
    print(name, bal)
    total = total + bal
  end
  return total
end

print("entries:", #entries)
print("trial balance total:", trial_balance())
