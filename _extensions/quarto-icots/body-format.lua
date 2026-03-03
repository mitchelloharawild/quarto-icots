-- body-format.lua
-- Pandoc/Quarto Lua filter that:
--   1. Inserts an empty BodyText paragraph before each heading
--   2. Indents the first line of each body paragraph by 720 twips (1.27 cm / 0.5")
--   3. Applies full (left + right) justification to body paragraphs
--   4. Inserts an empty BodyText paragraph before each list
--   5. Applies a 360-twip left indent to list-item paragraphs
--   6. Does not indent the abstract (it uses the Abstract style which has no indent)

------------------------------------------------------------------------
-- OpenXML snippets
------------------------------------------------------------------------

local BODY_PPR  = '<w:pPr><w:jc w:val="both"/><w:ind w:firstLine="720"/></w:pPr>'

local function empty_body_para()
  return pandoc.RawBlock(
    "openxml",
    '<w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr></w:p>'
  )
end

------------------------------------------------------------------------
-- Para handler for normal body paragraphs (indent + justify)
------------------------------------------------------------------------

local function indent_para(el)
  table.insert(el.content, 1,
    pandoc.RawInline("openxml", BODY_PPR))
  return el
end

------------------------------------------------------------------------
-- Returns true if the paragraph already has a RawInline that sets a
-- named paragraph style — used to detect the abstract paragraph, which
-- Pandoc/Quarto injects via a custom-style RawInline before we see it.
------------------------------------------------------------------------

local function has_pstyle(el, style)
  for _, inline in ipairs(el.content) do
    if inline.t == "RawInline"
      and inline.format == "openxml"
      and inline.text:find('w:pStyle w:val="' .. style .. '"', 1, true)
    then
      return true
    end
  end
  return false
end

------------------------------------------------------------------------
-- Walk a list (BulletList / OrderedList) and prepend a RawBlock pPr
-- override to every Plain/Para inside each item.  Pandoc merges a
-- RawBlock("openxml", "<w:pPr>…</w:pPr>") that immediately precedes a
-- Para/Plain into that paragraph's own <w:pPr>.
------------------------------------------------------------------------

local function fix_list_item_blocks(blocks)
  local out = pandoc.List()
  for _, block in ipairs(blocks) do
    if block.t == "Para" or block.t == "Plain" then
      out:insert(pandoc.RawBlock("openxml", "<w:pPr><w:ind w:left=\"360\"/></w:pPr>"))
      out:insert(block)
    else
      out:insert(block)
    end
  end
  return out
end

local function fix_list_items(list)
  if not FORMAT:match("docx") then return nil end

  for i, item in ipairs(list.content) do
    list.content[i] = fix_list_item_blocks(item)
  end

  return list
end

------------------------------------------------------------------------
-- Walk a Div that carries custom-style="Abstract" and apply
-- justify_para (no indent) to its paragraphs.
------------------------------------------------------------------------

local function justify_para(el)
  -- no-op: abstract paragraphs keep their own style, no extra pPr needed
  return el
end

local justify_only_filter = { Para = justify_para }

function Div(el)
  if not FORMAT:match("docx") then return nil end

  local cs = el.attributes["custom-style"]
  if cs == "Abstract" then
    return pandoc.walk_block(el, justify_only_filter)
  end
end

------------------------------------------------------------------------
-- Main Para filter — skip paragraphs that already carry the Abstract
-- style (injected by Pandoc from the YAML front-matter abstract field).
------------------------------------------------------------------------

function Para(el)
  if FORMAT:match("docx") then
    if has_pstyle(el, "Abstract") then
      return el
    end
    return indent_para(el)
  end
  return el
end

------------------------------------------------------------------------
-- Block-level pass: insert spacers before headings and lists
------------------------------------------------------------------------

function Blocks(blocks)
  if not FORMAT:match("docx") then return nil end

  local result = pandoc.List()
  for i, block in ipairs(blocks) do
    if block.t == "Header" and i > 1 then
      result:insert(empty_body_para())
    end
    if block.t == "BulletList" or block.t == "OrderedList" then
      result:insert(empty_body_para())
    end
    result:insert(block)
  end
  return result
end

------------------------------------------------------------------------
-- List handlers — must be assigned AFTER the functions are defined
------------------------------------------------------------------------

BulletList  = fix_list_items
OrderedList = fix_list_items

------------------------------------------------------------------------
-- Append a "References" H2 heading before the bibliography div
------------------------------------------------------------------------

function Pandoc(doc)
  if not FORMAT:match("docx") then return nil end

  -- Find an existing refs div and remove it so we can reinsert with heading
  local refs_div = nil
  local filtered = pandoc.List()
  for _, block in ipairs(doc.blocks) do
    if block.t == "Div" and block.identifier == "refs" then
      refs_div = block
    else
      filtered:insert(block)
    end
  end

  -- Append heading + refs div at the end
  filtered:insert(empty_body_para())
  filtered:insert(pandoc.Header(2, "References"))
  if refs_div then
    filtered:insert(refs_div)
  end

  doc.blocks = filtered
  return doc
end
