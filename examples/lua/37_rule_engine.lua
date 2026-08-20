-- App 37: Rule Engine for Fraud Checks

local rules = {
  {
    name = "large_amount",
    when_ = function(tx) return tx.amount > 10000 end,
    then_ = function(tx, out) out[#out + 1] = "manual_review" end,
  },
  {
    name = "cross_border",
    when_ = function(tx) return tx.country ~= tx.home_country end,
    then_ = function(tx, out) out[#out + 1] = "geo_check" end,
  },
  {
    name = "new_device",
    when_ = function(tx) return not tx.device_known end,
    then_ = function(tx, out) out[#out + 1] = "mfa_required" end,
  },
}

local txs = {
  {id = 1, amount = 1200, country = "US", home_country = "US", device_known = true},
  {id = 2, amount = 25000, country = "US", home_country = "US", device_known = false},
  {id = 3, amount = 700, country = "FR", home_country = "US", device_known = true},
}

for _, tx in ipairs(txs) do
  local actions = {}
  for _, rule in ipairs(rules) do
    if rule.when_(tx) then rule.then_(tx, actions) end
  end
  print("tx", tx.id, #actions > 0 and table.concat(actions, ",") or "approve")
end
