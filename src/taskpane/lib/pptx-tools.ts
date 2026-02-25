/**
 * PowerPoint parent tools — schemas and Office.js handler implementations.
 *
 * Each tool:
 *  1. Has a `schema` (ParentToolSchema) that is registered with the Usable embed.
 *  2. Has a `handler` that executes the Office.js operation and returns a plain object.
 *
 * All handlers wrap operations in PowerPoint.run() and call context.sync().
 * Read operations use .load() before sync and read values after.
 */

import { ParentToolSchema } from "./embed-sdk";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export type PptxToolHandler = (args: Record<string, unknown>) => Promise<unknown>;

export interface PptxTool {
  schema: ParentToolSchema;
  handler: PptxToolHandler;
}

// ---------------------------------------------------------------------------
// Tool definitions
// ---------------------------------------------------------------------------

const tools: PptxTool[] = [
  // ─── Presentation Info ────────────────────────────────────────────────────

  {
    schema: {
      name: "get_presentation_info",
      description:
        "Get high-level information about the current PowerPoint presentation: its title and total slide count.",
      parameters: {
        type: "object",
        properties: {},
      },
    },
    handler: async () => {
      return PowerPoint.run(async (context) => {
        const presentation = context.presentation;
        presentation.load("title");
        const slides = presentation.slides;
        slides.load("items");

        await context.sync();

        return {
          title: presentation.title,
          slideCount: slides.items.length,
        };
      });
    },
  },

  // ─── Slide Listing ────────────────────────────────────────────────────────

  {
    schema: {
      name: "list_slides",
      description:
        "List all slides in the presentation. Returns each slide's index (0-based) and ID.",
      parameters: {
        type: "object",
        properties: {},
      },
    },
    handler: async () => {
      return PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items/id");

        await context.sync();

        return {
          slides: slides.items.map((slide, index) => ({
            index,
            id: slide.id,
          })),
        };
      });
    },
  },

  // ─── Get Slide ────────────────────────────────────────────────────────────

  {
    schema: {
      name: "get_slide",
      description:
        "Get the contents of a specific slide by its 0-based index. Returns all shapes with their names, IDs, types, and text content.",
      parameters: {
        type: "object",
        properties: {
          index: {
            type: "number",
            description: "0-based index of the slide to read.",
          },
        },
        required: ["index"],
      },
    },
    handler: async (args) => {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(args.index as number);
        const shapes = slide.shapes;
        shapes.load("items/id,items/name,items/type");

        await context.sync();

        // Use getTextFrameOrNullObject() so shapes without a text frame
        // (images, charts, groups) return isNullObject=true instead of
        // throwing InvalidArgument at context.sync() time.
        const textFrames = shapes.items.map((shape) =>
          shape.getTextFrameOrNullObject()
        );
        textFrames.forEach((tf) => tf.load("isNullObject,textRange/text"));

        await context.sync();

        return {
          index: args.index as number,
          shapes: shapes.items.map((shape, i) => {
            const tf = textFrames[i];
            const text = tf.isNullObject ? null : tf.textRange.text;
            return {
              index: i,
              id: shape.id,
              name: shape.name,
              type: shape.type,
              text,
            };
          }),
        };
      });
    },
  },

  // ─── Get Selected Slide ───────────────────────────────────────────────────

  {
    schema: {
      name: "get_selected_slide",
      description:
        "Get the currently selected slide(s) in the presentation. Returns shapes and text on the first selected slide.",
      parameters: {
        type: "object",
        properties: {},
      },
    },
    handler: async () => {
      return PowerPoint.run(async (context) => {
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items/id");

        await context.sync();

        if (selectedSlides.items.length === 0) {
          return { selectedSlides: [] };
        }

        const firstSlide = selectedSlides.items[0];
        const shapes = firstSlide.shapes;
        shapes.load("items/id,items/name,items/type");

        await context.sync();

        const textFrames = shapes.items.map((shape) =>
          shape.getTextFrameOrNullObject()
        );
        textFrames.forEach((tf) => tf.load("isNullObject,textRange/text"));

        await context.sync();

        return {
          selectedSlideIds: selectedSlides.items.map((s) => s.id),
          firstSlide: {
            id: firstSlide.id,
            shapes: shapes.items.map((shape, i) => {
              const tf = textFrames[i];
              const text = tf.isNullObject ? null : tf.textRange.text;
              return {
                index: i,
                id: shape.id,
                name: shape.name,
                type: shape.type,
                text,
              };
            }),
          },
        };
      });
    },
  },

  // ─── Add Slide ────────────────────────────────────────────────────────────

  {
    schema: {
      name: "add_slide",
      description: "Add a new blank slide at the end of the presentation.",
      parameters: {
        type: "object",
        properties: {},
      },
    },
    handler: async () => {
      return PowerPoint.run(async (context) => {
        context.presentation.slides.add();

        await context.sync();

        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        return {
          success: true,
          newSlideCount: slides.items.length,
          newSlideIndex: slides.items.length - 1,
        };
      });
    },
  },

  // ─── Delete Slide ─────────────────────────────────────────────────────────

  {
    schema: {
      name: "delete_slide",
      description:
        "Delete a slide at the given 0-based index. This operation is permanent.",
      parameters: {
        type: "object",
        properties: {
          index: {
            type: "number",
            description: "0-based index of the slide to delete.",
          },
        },
        required: ["index"],
      },
    },
    handler: async (args) => {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(args.index as number);
        slide.delete();

        await context.sync();

        return { success: true };
      });
    },
  },

  // ─── Move Slide ───────────────────────────────────────────────────────────

  {
    schema: {
      name: "move_slide",
      description:
        "Move a slide to a new 0-based position in the presentation.",
      parameters: {
        type: "object",
        properties: {
          fromIndex: {
            type: "number",
            description: "0-based index of the slide to move.",
          },
          toIndex: {
            type: "number",
            description: "0-based destination index.",
          },
        },
        required: ["fromIndex", "toIndex"],
      },
    },
    handler: async (args) => {
      return PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        const fromIndex = args.fromIndex as number;
        const toIndex = args.toIndex as number;

        if (fromIndex < 0 || fromIndex >= slides.items.length) {
          throw new Error(`fromIndex ${fromIndex} is out of range (0–${slides.items.length - 1})`);
        }
        if (toIndex < 0 || toIndex >= slides.items.length) {
          throw new Error(`toIndex ${toIndex} is out of range (0–${slides.items.length - 1})`);
        }

        const slideToMove = slides.items[fromIndex];
        slideToMove.moveTo(toIndex);

        await context.sync();

        return { success: true, movedFrom: fromIndex, movedTo: toIndex };
      });
    },
  },

  // ─── Duplicate Slide ──────────────────────────────────────────────────────

  {
    schema: {
      name: "duplicate_slide",
      description:
        "Duplicate a slide by its 0-based index. The duplicate is inserted after the existing slides.",
      parameters: {
        type: "object",
        properties: {
          index: {
            type: "number",
            description: "0-based index of the slide to duplicate.",
          },
        },
        required: ["index"],
      },
    },
    handler: async (args) => {
      return PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items/id");
        await context.sync();

        const sourceSlide = slides.items[args.index as number];
        const sourceId = sourceSlide.id;

        // Export just this slide as base64, then re-insert it
        const base64Result = sourceSlide.exportAsBase64();
        await context.sync();

        context.presentation.insertSlidesFromBase64(base64Result.value, {
          sourceSlideIds: [sourceId],
          formatting: "UseDestinationTheme",
        });

        await context.sync();

        return { success: true, duplicatedSlideIndex: args.index as number };
      });
    },
  },

  // ─── Set Slide Title ──────────────────────────────────────────────────────

  {
    schema: {
      name: "set_slide_title",
      description:
        "Set the title text of a slide. Finds the title placeholder shape and sets its text.",
      parameters: {
        type: "object",
        properties: {
          index: {
            type: "number",
            description: "0-based index of the slide.",
          },
          title: {
            type: "string",
            description: "The title text to set.",
          },
        },
        required: ["index", "title"],
      },
    },
    handler: async (args) => {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(args.index as number);
        const shapes = slide.shapes;
        shapes.load("items/name,items/type,items/placeholderFormat");

        await context.sync();

        // Find the title placeholder by checking placeholderFormat.type
        let titleShapeIndex = -1;
        for (let i = 0; i < shapes.items.length; i++) {
          const shape = shapes.items[i];
          try {
            shape.placeholderFormat.load("type");
          } catch {
            // shape has no placeholder format
          }
        }

        await context.sync();

        for (let i = 0; i < shapes.items.length; i++) {
          const shape = shapes.items[i];
          try {
            const pType = shape.placeholderFormat.type;
            if (
              pType === PowerPoint.PlaceholderType.title ||
              pType === PowerPoint.PlaceholderType.centerTitle
            ) {
              titleShapeIndex = i;
              break;
            }
          } catch {
            // no placeholder format
          }
        }

        if (titleShapeIndex === -1) {
          // Fallback: use the first shape
          if (shapes.items.length === 0) {
            throw new Error(`Slide ${args.index} has no shapes to set as title.`);
          }
          titleShapeIndex = 0;
        }

        shapes.items[titleShapeIndex].textFrame.textRange.text = args.title as string;

        await context.sync();

        return { success: true };
      });
    },
  },

  // ─── Add Text Box ─────────────────────────────────────────────────────────

  {
    schema: {
      name: "add_text_box",
      description:
        "Add a text box to a slide at the specified position and size (in points).",
      parameters: {
        type: "object",
        properties: {
          index: {
            type: "number",
            description: "0-based index of the slide.",
          },
          text: {
            type: "string",
            description: "Text content for the text box.",
          },
          left: {
            type: "number",
            description: "Horizontal position from the left edge in points. Default: 50.",
          },
          top: {
            type: "number",
            description: "Vertical position from the top edge in points. Default: 50.",
          },
          width: {
            type: "number",
            description: "Width of the text box in points. Default: 400.",
          },
          height: {
            type: "number",
            description: "Height of the text box in points. Default: 100.",
          },
        },
        required: ["index", "text"],
      },
    },
    handler: async (args) => {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(args.index as number);

        const left   = (args.left   as number | undefined) ?? 50;
        const top    = (args.top    as number | undefined) ?? 50;
        const width  = (args.width  as number | undefined) ?? 400;
        const height = (args.height as number | undefined) ?? 100;

        slide.shapes.addTextBox(args.text as string, {
          left,
          top,
          width,
          height,
        });

        await context.sync();

        return { success: true };
      });
    },
  },

  // ─── Set Shape Text ───────────────────────────────────────────────────────

  {
    schema: {
      name: "set_shape_text",
      description:
        "Set the text content of a shape on a slide, identified by its 0-based shape index.",
      parameters: {
        type: "object",
        properties: {
          slideIndex: {
            type: "number",
            description: "0-based index of the slide.",
          },
          shapeIndex: {
            type: "number",
            description: "0-based index of the shape within the slide.",
          },
          text: {
            type: "string",
            description: "Text to set on the shape.",
          },
        },
        required: ["slideIndex", "shapeIndex", "text"],
      },
    },
    handler: async (args) => {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(args.slideIndex as number);
        const shape = slide.shapes.getItemAt(args.shapeIndex as number);
        shape.textFrame.textRange.text = args.text as string;

        await context.sync();

        return { success: true };
      });
    },
  },

  // ─── Get Shapes ───────────────────────────────────────────────────────────

  {
    schema: {
      name: "get_shapes",
      description:
        "Get all shapes on a slide with their names, IDs, and types.",
      parameters: {
        type: "object",
        properties: {
          index: {
            type: "number",
            description: "0-based index of the slide.",
          },
        },
        required: ["index"],
      },
    },
    handler: async (args) => {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(args.index as number);
        const shapes = slide.shapes;
        shapes.load("items/id,items/name,items/type");

        await context.sync();

        return {
          slideIndex: args.index as number,
          shapes: shapes.items.map((shape, i) => ({
            index: i,
            id: shape.id,
            name: shape.name,
            type: shape.type,
          })),
        };
      });
    },
  },

  // ─── Delete Shape ─────────────────────────────────────────────────────────

  {
    schema: {
      name: "delete_shape",
      description:
        "Delete a shape from a slide by its 0-based shape index.",
      parameters: {
        type: "object",
        properties: {
          slideIndex: {
            type: "number",
            description: "0-based index of the slide.",
          },
          shapeIndex: {
            type: "number",
            description: "0-based index of the shape to delete.",
          },
        },
        required: ["slideIndex", "shapeIndex"],
      },
    },
    handler: async (args) => {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(args.slideIndex as number);
        const shape = slide.shapes.getItemAt(args.shapeIndex as number);
        shape.delete();

        await context.sync();

        return { success: true };
      });
    },
  },

  // ─── Add Table ────────────────────────────────────────────────────────────

  {
    schema: {
      name: "add_table",
      description:
        "Add a table to a slide with the specified number of rows and columns.",
      parameters: {
        type: "object",
        properties: {
          index: {
            type: "number",
            description: "0-based index of the slide.",
          },
          rows: {
            type: "number",
            description: "Number of rows in the table.",
          },
          columns: {
            type: "number",
            description: "Number of columns in the table.",
          },
          left: {
            type: "number",
            description: "Horizontal position in points. Default: 50.",
          },
          top: {
            type: "number",
            description: "Vertical position in points. Default: 100.",
          },
          width: {
            type: "number",
            description: "Width of the table in points. Default: 500.",
          },
          height: {
            type: "number",
            description: "Height of the table in points. Default: 200.",
          },
        },
        required: ["index", "rows", "columns"],
      },
    },
    handler: async (args) => {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(args.index as number);

        const left    = (args.left    as number | undefined) ?? 50;
        const top     = (args.top     as number | undefined) ?? 100;
        const width   = (args.width   as number | undefined) ?? 500;
        const height  = (args.height  as number | undefined) ?? 200;

        slide.shapes.addTable(
          args.rows as number,
          args.columns as number,
          { left, top, width, height }
        );

        await context.sync();

        return { success: true };
      });
    },
  },

  // ─── Set Table Data ───────────────────────────────────────────────────────

  {
    schema: {
      name: "set_table_data",
      description:
        "Set the text content of a specific table cell on a slide. The table is identified by its shape index.",
      parameters: {
        type: "object",
        properties: {
          slideIndex: {
            type: "number",
            description: "0-based index of the slide.",
          },
          shapeIndex: {
            type: "number",
            description: "0-based index of the table shape on the slide.",
          },
          row: {
            type: "number",
            description: "0-based row index.",
          },
          column: {
            type: "number",
            description: "0-based column index.",
          },
          text: {
            type: "string",
            description: "Text to set in the cell.",
          },
        },
        required: ["slideIndex", "shapeIndex", "row", "column", "text"],
      },
    },
    handler: async (args) => {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(args.slideIndex as number);
        const tableShape = slide.shapes.getItemAt(args.shapeIndex as number);
        const table = tableShape.getTable();

        const cell = table.getCellOrNullObject(
          args.row as number,
          args.column as number
        );

        cell.text = args.text as string;

        await context.sync();

        return { success: true };
      });
    },
  },

  // ─── Apply Background Color ───────────────────────────────────────────────

  {
    schema: {
      name: "apply_background_color",
      description:
        "Apply a solid background color to a slide.",
      parameters: {
        type: "object",
        properties: {
          index: {
            type: "number",
            description: "0-based index of the slide.",
          },
          color: {
            type: "string",
            description: "Hex color string, e.g. '#FF5733' or '#FFFFFF'.",
          },
        },
        required: ["index", "color"],
      },
    },
    handler: async (args) => {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(args.index as number);
        slide.background.fill.setSolidFill({ color: args.color as string });

        await context.sync();

        return { success: true };
      });
    },
  },

  // ─── Set Shape Position ───────────────────────────────────────────────────

  {
    schema: {
      name: "set_shape_position",
      description:
        "Move and/or resize a shape on a slide by setting its left, top, width, and/or height (in points). All position/size properties are optional — supply only the ones you want to change.",
      parameters: {
        type: "object",
        properties: {
          slideIndex: {
            type: "number",
            description: "0-based index of the slide.",
          },
          shapeIndex: {
            type: "number",
            description: "0-based index of the shape on the slide.",
          },
          left: {
            type: "number",
            description: "Distance from the left edge of the slide in points.",
          },
          top: {
            type: "number",
            description: "Distance from the top edge of the slide in points.",
          },
          width: {
            type: "number",
            description: "Width of the shape in points.",
          },
          height: {
            type: "number",
            description: "Height of the shape in points.",
          },
        },
        required: ["slideIndex", "shapeIndex"],
      },
    },
    handler: async (args) => {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(args.slideIndex as number);
        const shape = slide.shapes.getItemAt(args.shapeIndex as number);
        shape.load("left,top,width,height");
        await context.sync();

        if (args.left !== undefined) shape.left = args.left as number;
        if (args.top !== undefined) shape.top = args.top as number;
        if (args.width !== undefined) shape.width = args.width as number;
        if (args.height !== undefined) shape.height = args.height as number;

        await context.sync();

        return {
          success: true,
          left: shape.left,
          top: shape.top,
          width: shape.width,
          height: shape.height,
        };
      });
    },
  },

  // ─── Set Layout ───────────────────────────────────────────────────────────

  {
    schema: {
      name: "set_layout",
      description:
        "Apply a slide layout to a slide. Layouts are found on the first slide master. Use list_layouts conceptually — common names include 'Blank', 'Title Slide', 'Title and Content'.",
      parameters: {
        type: "object",
        properties: {
          index: {
            type: "number",
            description: "0-based index of the slide.",
          },
          layoutIndex: {
            type: "number",
            description: "0-based index of the layout in the slide master's layout collection.",
          },
        },
        required: ["index", "layoutIndex"],
      },
    },
    handler: async (args) => {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(args.index as number);
        const slideMasters = context.presentation.slideMasters;
        slideMasters.load("items");
        await context.sync();

        if (slideMasters.items.length === 0) {
          throw new Error("No slide masters found in this presentation.");
        }

        const masterLayouts = slideMasters.items[0].layouts;
        masterLayouts.load("items");
        await context.sync();

        const layoutIdx = args.layoutIndex as number;
        if (layoutIdx < 0 || layoutIdx >= masterLayouts.items.length) {
          throw new Error(
            `layoutIndex ${layoutIdx} is out of range (0–${masterLayouts.items.length - 1}). Available layouts: ${masterLayouts.items.length}`
          );
        }

        const layout = masterLayouts.items[layoutIdx];
        slide.applyLayout(layout);

        await context.sync();

        return { success: true };
      });
    },
  },
];

// ---------------------------------------------------------------------------
// Exports
// ---------------------------------------------------------------------------

export const pptxToolSchemas: ParentToolSchema[] = tools.map((t) => t.schema);

const handlerMap: Map<string, PptxToolHandler> = new Map(
  tools.map((t) => [t.schema.name, t.handler])
);

// WHY BOTH OF THESE EXIST
// ─────────────────────────────────────────────────────────────────────────────
//
// Problem 1 — Double-mount deduplication (inFlightCalls)
// ───────────────────────────────────────────────────────
// In this environment (Office.js task pane on macOS / WKWebView + React legacy
// mode), two `window.addEventListener("message", ...)` listeners can become
// active simultaneously for a single UsableChatEmbed instance. This is caused
// by a combination of:
//
//   1. React legacy mode (ReactDOM.render) does not batch async state updates.
//      The two setState calls in useAuth's refreshAccessToken fire as separate
//      renders, which can briefly double-mount ChatPane and leave an orphaned
//      listener from the first mount.
//
//   2. Office.js / WKWebView has known quirks where the task pane JS context
//      can initialise in ways that React's effect cleanup does not fully
//      intercept.
//
// The result: a single TOOL_CALL postMessage is received by both listeners,
// causing handlePptxToolCall — and therefore PowerPoint.run() — to execute twice.
// Non-idempotent operations like add_slide produce duplicate artefacts;
// others silently double-write or fail on the second attempt.
//
// FIX: track in-flight calls by (toolName + serialised args). If an identical
// call arrives while one is already running, return the same promise so
// PowerPoint.run() only executes once. The second TOOL_RESPONSE is harmless.
const inFlightCalls = new Map<string, Promise<unknown>>();

// Problem 2 — Concurrent PowerPoint.run() calls (executionChain)
// ───────────────────────────────────────────────────────────────
// When the AI sends many tool calls in rapid succession (e.g. 14 set_table_data
// calls to populate a table), the embed SDK dispatches them all immediately.
// Each dispatch triggers a separate PowerPoint.run() context. PowerPoint.js
// cannot handle many concurrent contexts — some never resolve their promise,
// producing a `pending_parent_response` on the AI side.
//
// FIX: serialize all PowerPoint.run() calls through a single promise chain.
// Each call waits for the previous one to settle before starting. Errors on
// one call are swallowed from the chain so they don't block subsequent calls.
let executionChain: Promise<unknown> = Promise.resolve();

function serializeExecution<T>(fn: () => Promise<T>): Promise<T> {
  const result = executionChain.then(() => fn()) as Promise<T>;
  executionChain = result.then(
    () => undefined,
    () => undefined, // don't let a failed call block the queue
  );
  return result;
}

export async function handlePptxToolCall(
  toolName: string,
  args: unknown
): Promise<unknown> {
  const handler = handlerMap.get(toolName);
  if (!handler) {
    throw new Error(`Unknown PowerPoint tool: "${toolName}"`);
  }

  const key = `${toolName}:${JSON.stringify(args)}`;
  const existing = inFlightCalls.get(key);
  if (existing) return existing;

  const promise = serializeExecution(() => handler(args as Record<string, unknown>)).then(
    (result) => { inFlightCalls.delete(key); return result; },
    (err)    => { inFlightCalls.delete(key); throw err; }
  );
  inFlightCalls.set(key, promise);
  return promise;
}
