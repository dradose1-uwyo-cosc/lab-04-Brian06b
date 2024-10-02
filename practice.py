items = ["milk", "clothes", "tv", "Lego", "hat"]
prices = [2, 20, 79, 12, 14]

# I bought Eggs for $3.4

max_price = max(prices)
print(f"{max_price}")

new_items = items.remove(items[2])
new_prices = prices.remove(max_price)
print(items)
print(prices)