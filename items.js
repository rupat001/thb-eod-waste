// ============================================================
//  THE ITEM LIST — the only place items are defined.
//  Staff form, manager page and the Google Sheet all follow it.
//
//  Each item: [key, label, note]
//   - key:   never change it once used (the Sheet column is named after it).
//            lowercase letters, numbers and _ only.
//   - label: what staff see. Safe to rename any time.
//   - note:  optional small text under the label (e.g. "scoops").
//
//  ADD an item:    add a line in the right group. The Sheet gets a new
//                  column automatically on the next save.
//  REMOVE an item: delete its line. Its old numbers stay in the Sheet.
//  REORDER:        move lines. Only changes what staff see.
// ============================================================
window.EOD_ITEMS = [
  ["Donuts", [
    ["apple_fritter", "Apple Fritter"],
    ["oreo", "Oreo"],
    ["vanilla_cream", "Vanilla Cream"],
    ["boston_cream", "Boston Cream"],
    ["canadian_maple", "Canadian Maple"],
    ["raspberry", "Raspberry"],
    ["vanilla_dip", "Vanilla Dip"],
    ["chocolate_dip", "Chocolate Dip"],
    ["maple_dip", "Maple Dip"],
    ["honey_dip", "Honey Dip"],
    ["honey_cruller", "Honey Cruller"],
    ["chocolate_cruller", "Chocolate Cruller"],
    ["sour_cream", "Sour Cream"],
    ["blueberry_sour_cream", "Blueberry Sour Cream"],
    ["chocolate_glazed", "Chocolate Glazed"],
    ["double_chocolate", "Double Chocolate"],
    ["old_fashioned_plain", "Old Fashioned Plain"],
    ["yeast_rings", "Yeast Rings"],
    ["yeast_fills", "Yeast Fills"]
  ]],
  ["Timbits", [
    ["tb_honey_dip", "Honey Dip"],
    ["tb_chocolate_glazed", "Chocolate Glazed"],
    ["tb_blueberry_sour_cream", "Blueberry Sour Cream"],
    ["tb_birthday_cake", "Birthday Cake"]
  ]],
  ["Cookies", [
    ["chocolate_chunk", "Chocolate Chunk"],
    ["salted_caramel", "Salted Caramel"]
  ]],
  ["Muffins", [
    ["muffin_chocolate_chip", "Chocolate Chip"],
    ["muffin_blueberry", "Blueberry"]
  ]],
  ["Sandwich", [
    ["eggs", "Eggs"],
    ["hashbrown", "Hashbrown"],
    ["sausage", "Sausage"],
    ["bacon_egg_bites", "Bacon Egg Bites"],
    ["spinach_egg_bites", "Spinach Egg Bites"],
    ["bacon", "Bacon", "slices, as counted"],
    ["chicken", "Chicken", "scoops"],
    ["steak", "Steak", "scoops"]
  ]],
  ["Bagels", [
    ["bagel_plain", "Plain"],
    ["bagel_everything", "Everything"]
  ]],
  ["Breads", [
    ["biscuit", "Biscuit"],
    ["english_muffin", "English Muffin"],
    ["croissant", "Croissant"],
    ["sourdough", "Sourdough Bread"],
    ["artisan_bun", "Artisan Bun"]
  ]]
];

// Every key, in list order (sent with saves so the Sheet adds new columns in order).
window.EOD_ITEM_KEYS = window.EOD_ITEMS.flatMap(function (g) { return g[1].map(function (i) { return i[0]; }); });
