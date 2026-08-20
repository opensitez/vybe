-- Full App 05: Payroll Engine with tax brackets and overtime

local employees = {
  {name = "Nora", hours = 168, rate = 42.5, bonuses = {200, 100}},
  {name = "Milo", hours = 154, rate = 35.0, bonuses = {75}},
  {name = "Ari", hours = 182, rate = 55.0, bonuses = {500, 250}},
}

local tax_brackets = {
  {up_to = 3000, rate = 0.10},
  {up_to = 6000, rate = 0.20},
  {up_to = math.huge, rate = 0.30},
}

local function progressive_tax(income)
  local tax, prev = 0, 0
  for _, b in ipairs(tax_brackets) do
    if income > prev then
      local taxable = math.min(income, b.up_to) - prev
      tax = tax + taxable * b.rate
      prev = b.up_to
    end
  end
  return tax
end

for _, e in ipairs(employees) do
  local overtime = math.max(0, e.hours - 160)
  local base_hours = e.hours - overtime
  local gross = base_hours * e.rate + overtime * e.rate * 1.5
  for _, b in ipairs(e.bonuses) do gross = gross + b end
  local tax = progressive_tax(gross)
  local net = gross - tax
  print(string.format("%s gross=%.2f tax=%.2f net=%.2f overtime=%dh", e.name, gross, tax, net, overtime))
end
