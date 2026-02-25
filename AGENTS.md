# Agent Guide — Usable PowerPoint Tools

## Shape Position Data

`get_slide`, `get_selected_slide`, and `get_shapes` return **position and size** for every shape. All values are in **points** (Office.js unit; 72 points = 1 inch).

### Shape object shape

```json
{
  "index": 0,
  "id": "some-id",
  "name": "Title 1",
  "type": "GeometricShape",
  "left": 57.6,
  "top": 38.4,
  "width": 828.0,
  "height": 130.0,
  "text": "My slide title"
}
```

`text` is `null` for shapes that have no text frame (images, charts, etc.).
`get_shapes` does not include `text` — use `get_slide` or `get_selected_slide` when you need both position and text in one call.

---

## Targeting shapes by ID vs index

**Always prefer `shapeId`** over `shapeIndex` when calling `delete_shape` or `set_shape_position`. Indices shift after every deletion; IDs are stable.

```json
{ "slideIndex": 0, "shapeId": "some-id", "left": 57.6, "top": 200.0 }
```

---

## Recommended workflow for layout fixes

1. Call `get_slide` (or `get_selected_slide`) to get all shapes with positions, IDs, and text.
2. Identify the specific shape(s) to change from the data.
3. Call `set_shape_position` or `delete_shape` for each affected shape by **id**.
4. Call `get_slide` again to confirm the result.

---

## Duplicate detection — read carefully

**Only treat shapes as duplicates when ALL of the following are true:**
- They have **identical non-empty `text` content**, AND
- Their `left` and `top` coordinates are **within 1 point of each other**

**Do NOT use `name` as a duplicate signal.** PowerPoint auto-names every new text box sequentially ("TextBox 1", "TextBox 2", …). Similar names are completely normal and expected — they do not indicate duplicates.

**When in doubt, ask the user before deleting anything.** A wrong deletion is far more disruptive than pausing to confirm.

---

## Clearing a slide

Do **not** call `delete_all_shapes` unless the user explicitly asks to wipe and rebuild the slide. For targeted fixes (remove one shape, reposition another), use `delete_shape` or `set_shape_position` by `shapeId`.

---

## Bounding box overlap check

Two shapes overlap when:
```
overlap = (a.left < b.left + b.width)  && (a.left + a.width > b.left)
       && (a.top  < b.top  + b.height) && (a.top  + a.height > b.top)
```

---

## Slide dimensions (widescreen default)

- Width: **960 pt**
- Height: **540 pt**

Use these as reference when positioning or sizing shapes.
