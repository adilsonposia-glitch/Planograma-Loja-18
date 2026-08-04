---
name: Lumina Spectral
colors:
  surface: '#0b1326'
  surface-dim: '#0b1326'
  surface-bright: '#31394d'
  surface-container-lowest: '#060e20'
  surface-container-low: '#131b2e'
  surface-container: '#171f33'
  surface-container-high: '#222a3d'
  surface-container-highest: '#2d3449'
  on-surface: '#dae2fd'
  on-surface-variant: '#b9cacb'
  inverse-surface: '#dae2fd'
  inverse-on-surface: '#283044'
  outline: '#849495'
  outline-variant: '#3b494b'
  surface-tint: '#00dbe9'
  primary: '#dbfcff'
  on-primary: '#00363a'
  primary-container: '#00f0ff'
  on-primary-container: '#006970'
  inverse-primary: '#006970'
  secondary: '#c0c1ff'
  on-secondary: '#1000a9'
  secondary-container: '#3131c0'
  on-secondary-container: '#b0b2ff'
  tertiary: '#fff3f2'
  on-tertiary: '#67001b'
  tertiary-container: '#ffcdcf'
  on-tertiary-container: '#bc0b3b'
  error: '#ffb4ab'
  on-error: '#690005'
  error-container: '#93000a'
  on-error-container: '#ffdad6'
  primary-fixed: '#7df4ff'
  primary-fixed-dim: '#00dbe9'
  on-primary-fixed: '#002022'
  on-primary-fixed-variant: '#004f54'
  secondary-fixed: '#e1e0ff'
  secondary-fixed-dim: '#c0c1ff'
  on-secondary-fixed: '#07006c'
  on-secondary-fixed-variant: '#2f2ebe'
  tertiary-fixed: '#ffdadb'
  tertiary-fixed-dim: '#ffb2b7'
  on-tertiary-fixed: '#40000d'
  on-tertiary-fixed-variant: '#92002a'
  background: '#0b1326'
  on-background: '#dae2fd'
  surface-variant: '#2d3449'
typography:
  headline-lg:
    fontFamily: Geist
    fontSize: 32px
    fontWeight: '600'
    lineHeight: 40px
    letterSpacing: -0.02em
  headline-md:
    fontFamily: Geist
    fontSize: 24px
    fontWeight: '600'
    lineHeight: 32px
    letterSpacing: -0.01em
  headline-sm:
    fontFamily: Geist
    fontSize: 20px
    fontWeight: '500'
    lineHeight: 28px
  body-lg:
    fontFamily: Geist
    fontSize: 16px
    fontWeight: '400'
    lineHeight: 24px
  body-md:
    fontFamily: Geist
    fontSize: 14px
    fontWeight: '400'
    lineHeight: 20px
  data-display:
    fontFamily: JetBrains Mono
    fontSize: 18px
    fontWeight: '500'
    lineHeight: 24px
    letterSpacing: -0.03em
  label-caps:
    fontFamily: JetBrains Mono
    fontSize: 11px
    fontWeight: '700'
    lineHeight: 16px
    letterSpacing: 0.08em
  label-sm:
    fontFamily: Geist
    fontSize: 12px
    fontWeight: '500'
    lineHeight: 16px
rounded:
  sm: 0.125rem
  DEFAULT: 0.25rem
  md: 0.375rem
  lg: 0.5rem
  xl: 0.75rem
  full: 9999px
spacing:
  unit: 4px
  gutter: 16px
  margin: 24px
  container-padding: 20px
  panel-gap: 12px
---

## Brand & Style
The design system focuses on high-density data visualization and real-time operational monitoring for modern retail environments. The brand personality is technical, precise, and authoritative, designed to instill confidence in large-scale logistical decision-making.

The aesthetic follows a **Modern Corporate** approach with **Glassmorphic** accents to handle complex data overlays without losing spatial context. The UI utilizes deep, light-absorbing backgrounds to maximize the legibility of vibrant "spectral" heatmaps. This creates a "mission control" atmosphere where critical stock or compliance alerts command immediate attention against a subdued, architectural backdrop.

## Colors
The palette is engineered for a dark-room "command center" environment. The primary color is a high-visibility Cyan used for interactive elements and data highlights. 

- **Surface System:** The foundation uses a deep Navy (`#0F172A`) for the base layer, with Slate transitions for containers to maintain depth.
- **Heatmap Palette:** A strictly defined spectral scale (Green-Yellow-Red) is used exclusively for compliance and stock levels. 
    - **Optimal (#10B981):** High stock, 95%+ compliance.
    - **Warning (#F59E0B):** Low stock, approaching non-compliance.
    - **Critical (#EF4444):** Out of stock, immediate action required.
- **Data Accents:** Secondary Indigo is used for non-critical trend lines and historical comparisons.

## Typography
The typography system prioritizes technical legibility and information density. 

- **Primary Typeface:** **Geist** is used for all UI labels, headings, and standard body text. Its clean, geometric sans-serif nature feels modern and unobtrusive.
- **Data Typeface:** **JetBrains Mono** is employed for numeric data, coordinates, stock counts, and timestamps. The monospaced nature ensures that columns of numbers remain perfectly aligned during real-time updates.
- **Hierarchy:** Use `label-caps` for table headers and category labels over maps. `data-display` should be the default for KPI cards and inventory counts.

## Layout & Spacing
The layout uses a **Fluid Grid** system designed to maximize screen real estate for map-based analytics. 

- **Structure:** A persistent left-hand "control rail" (64px collapsed / 240px expanded) handles high-level navigation. The main content area uses a 12-column grid.
- **Density:** Spacing is tight (4px base unit) to allow for the display of large data arrays without excessive scrolling. 
- **Adaptation:** On mobile/tablet, the side panels transform into bottom sheets to keep the central map or heatmap view unobstructed. Panels should utilize a "stacking" logic where secondary data hides behind primary tabs.

## Elevation & Depth
Depth is conveyed through **Tonal Layers** and **Glassmorphism**, avoiding traditional shadows to keep the interface looking crisp and digital.

- **Base Layer:** The darkest navy (`#020617`), reserved for the global background.
- **Surface Layer:** `rgba(30, 41, 59, 0.7)` with a 12px backdrop blur. This is used for floating panels, tooltips, and map overlays.
- **Stroke:** All containers must have a 1px solid border (`rgba(255, 255, 255, 0.1)`) to define boundaries against the dark background.
- **Interactive States:** Hovering over a card should increase its border opacity rather than adding a shadow, maintaining the "light-emitting" feel of the dashboard.

## Shapes
This design system utilizes a **Soft (1)** shape language. The 0.25rem (4px) base radius provides a modern, technical feel without the clinical harshness of sharp corners.

- **Standard Elements:** Inputs, buttons, and small tags use a 4px radius.
- **Large Containers:** Cards and map overlays use `rounded-lg` (8px) to subtly differentiate them from the smaller UI controls.
- **Data Points:** Heatmap nodes on the store map are circular (pill-shaped) to represent signal dispersion.

## Components
- **Buttons:** Primary buttons are solid Cyan with black text for maximum contrast. Secondary buttons use a ghost style (border only).
- **KPI Cards:** Feature a large `data-display` number. A small trend sparkline should be positioned in the bottom half of the card.
- **Data Chips:** Small, high-contrast badges for status (e.g., "In Stock"). These use the heatmap color palette for the background with a 10% opacity tint and 100% opacity text.
- **Input Fields:** Dark backgrounds (`#1E293B`) with a 1px border. The border glows Cyan when the field is focused.
- **Store Map Overlays:** Translucent panels with "Geist" labels. Use high-contrast white text for data points placed directly over red or green heatmap zones to ensure legibility.
- **Lists:** Data tables should have no horizontal borders, only subtle vertical dividers between key data columns to emphasize vertical scanning of metrics.