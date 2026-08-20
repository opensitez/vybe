-- App 21: Ticket Booking System

local events = {
  {id = 1, name = "LuaConf Workshop", seats = 5, price = 120},
  {id = 2, name = "Game Jam Night", seats = 3, price = 45},
}

local bookings = {}

local function find_event(id)
  for _, e in ipairs(events) do
    if e.id == id then return e end
  end
end

local function reserve(user, event_id, seats)
  local e = find_event(event_id)
  if not e then return false, "event not found" end
  if seats <= 0 then return false, "invalid seat count" end
  if e.seats < seats then return false, "not enough seats" end

  e.seats = e.seats - seats
  local order = {
    order_id = #bookings + 1,
    user = user,
    event = e.name,
    seats = seats,
    total = seats * e.price,
    status = "confirmed"
  }
  bookings[#bookings + 1] = order
  return true, order
end

local function cancel(order_id)
  for _, o in ipairs(bookings) do
    if o.order_id == order_id and o.status == "confirmed" then
      o.status = "cancelled"
      for _, e in ipairs(events) do
        if e.name == o.event then e.seats = e.seats + o.seats end
      end
      return true
    end
  end
  return false
end

local ok1, o1 = reserve("alice", 1, 2)
local ok2, o2 = reserve("bob", 2, 1)
print("reserve alice:", ok1, ok1 and o1.total or o1)
print("reserve bob:", ok2, ok2 and o2.total or o2)
print("cancel #1:", cancel(1))

for _, e in ipairs(events) do
  print(string.format("event=%s seats_left=%d", e.name, e.seats))
end
