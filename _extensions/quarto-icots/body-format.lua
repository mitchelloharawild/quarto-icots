-- body-format.lua
-- Pandoc/Quarto Lua filter that:
--   1. Inserts an empty BodyText paragraph before each heading
--   2. Indents the first line of each body paragraph by 720 twips (1.27 cm / 0.5")
--   3. Applies full (left + right) justification to body paragraphs
--   4. Inserts an empty BodyText paragraph before each top-level list
--   5. Applies a depth-aware left indent to list-item paragraphs
--      (360 twips per level: depth 0 → 360, depth 1 → 720, depth 2 → 1080, …)
--   6. Does not indent the abstract (it uses the Abstract style which has no indent)
--
-- All structural transformation is done in the Pandoc (document-level) filter
-- so that we walk the block tree manually and Pandoc never auto-traverses into
-- nodes we have already processed.

------------------------------------------------------------------------
-- OpenXML helpers
------------------------------------------------------------------------

local BODY_PPR = '<w:pPr><w:jc w:val="both"/><w:ind w:firstLine="720"/></w:pPr>'

local function empty_body_para()
  return pandoc.RawBlock(
    "openxml",
    '<w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr></w:p>'
  )
end

local function list_ppr(depth)
  return pandoc.RawBlock(
    "openxml",
    string.format('<w:pPr><w:ind w:left="%d"/></w:pPr>', 360 * (depth + 1))
  )
end

------------------------------------------------------------------------
-- Returns true when a Para/Plain already carries a named pStyle
-- (e.g. "Abstract") injected by Pandoc as a RawInline.
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
-- Forward declaration for mutual recursion
------------------------------------------------------------------------

local process_blocks  -- processes a flat block list, returns a new pandoc.List

------------------------------------------------------------------------
-- Process the blocks belonging to one list item at a given depth.
-- Para/Plain nodes get a depth-aware pPr prepended.
-- Nested lists are handled recursively (no spacer para for nested lists).
-- Other block types are recursed into via process_blocks.
------------------------------------------------------------------------

local function process_item_blocks(blocks, depth)
  local out = pandoc.List()
  for _, block in ipairs(blocks) do
    if block.t == "Para" or block.t == "Plain" then
      out:insert(list_ppr(depth))
      out:insert(block)
    elseif block.t == "BulletList" or block.t == "OrderedList" then
      -- nested list: no spacer, just recurse one level deeper
      for i, item in ipairs(block.content) do
        block.content[i] = process_item_blocks(item, depth + 1)
      end
      out:insert(block)
    else
      -- e.g. BlockQuote, Table inside a list item — recurse generically
      out:extend(process_blocks({ block }, false))
    end
  end
  return out
end

------------------------------------------------------------------------
-- Walk a flat list of blocks, returning a new list.
-- top_level=true  → insert spacer before lists and headings (after block 1)
-- top_level=false → no spacers (used when recursing into non-list containers)
------------------------------------------------------------------------

process_blocks = function(blocks, top_level)
  local out = pandoc.List()
  for i, block in ipairs(blocks) do
    if block.t == "Header" then
      if top_level and i > 1 then
        out:insert(empty_body_para())
      end
      out:insert(block)

    elseif block.t == "BulletList" or block.t == "OrderedList" then
      if top_level then
        out:insert(empty_body_para())
      end
      for j, item in ipairs(block.content) do
        block.content[j] = process_item_blocks(item, 0)
      end
      out:insert(block)

    elseif block.t == "Para" or block.t == "Plain" then
      -- Apply body indent/justify unless it already carries a named style
      if not has_pstyle(block, "Abstract") then
        table.insert(block.content, 1,
          pandoc.RawInline("openxml", BODY_PPR))
      end
      out:insert(block)

    elseif block.t == "Div" then
      -- Recurse into Divs (e.g. the refs div, abstract div)
      block.content = process_blocks(block.content, false)
      out:insert(block)

    elseif block.t == "BlockQuote" then
      block.content = process_blocks(block.content, false)
      out:insert(block)

    else
      out:insert(block)
    end
  end
  return out
end

------------------------------------------------------------------------
-- Document-level filter: single pass over the whole block tree.
-- Pandoc's auto-traversal fires BEFORE Pandoc(), so we rely solely on
-- the manual walk above — no BulletList / Para / Blocks filters are
-- registered, avoiding all double-processing.
------------------------------------------------------------------------

function Pandoc(doc)
  if not FORMAT:match("docx") then return nil end

  -- Separate out the refs div so we can append it after a heading
  local refs_div = nil
  local main = pandoc.List()
  for _, block in ipairs(doc.blocks) do
    if block.t == "Div" and block.identifier == "refs" then
      refs_div = block
    else
      main:insert(block)
    end
  end

  local result = process_blocks(main, true)

  result:insert(empty_body_para())
  result:insert(pandoc.Header(2, "References"))
  if refs_div then
    result:insert(refs_div)
  end

  doc.blocks = result
  return doc
end
