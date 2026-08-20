-- Full App 01: Banking Ledger and Monthly Statement

local accounts = {
  alice = {balance = 1200, tx = {}},
  bob = {balance = 600, tx = {}},
  ops = {balance = 10000, tx = {}},
}

local function post(name, kind, amount, note)
  local a = accounts[name]
  assert(a, "unknown account")
  if kind == "debit" then
    assert(a.balance >= amount, "insufficient funds")
    a.balance = a.balance - amount
  elseif kind == "credit" then
    a.balance = a.balance + amount
  else
    error("unknown kind")
  end
  a.tx[#a.tx + 1] = {kind = kind, amount = amount, note = note}
end

local function transfer(from, to, amount, note)
  post(from, "debit", amount, "to " .. to .. " | " .. note)
  post(to, "credit", amount, "from " .. from .. " | " .. note)
end

transfer("alice", "bob", 150, "rent split")
post("alice", "debit", 80, "groceries")
post("bob", "credit", 400, "freelance")
transfer("ops", "alice", 1200, "salary")

for name, a in pairs(accounts) do
  local debit, credit = 0, 0
  for _, t in ipairs(a.tx) do
    if t.kind == "debit" then debit = debit + t.amount else credit = credit + t.amount end
  end
  print(string.format("%s balance=%0.2f credit=%0.2f debit=%0.2f", name, a.balance, credit, debit))
  for i, t in ipairs(a.tx) do
    print(string.format("  %02d %-6s %7.2f %s", i, t.kind, t.amount, t.note))
  end
end
