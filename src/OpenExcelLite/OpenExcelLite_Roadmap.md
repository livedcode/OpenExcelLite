
# OpenExcelLite — Future Roadmap (Personal Reference)

This document lists the planned features, enhancements, and long‑term goals for OpenExcelLite.

---

# 🚀 PHASE 1 — Core Enhancements (High Priority)

These improvements will strengthen the core functionality without increasing library size significantly.

### 1. **Multiple Streaming Sheets**
- Support more than one streaming worksheet per workbook.
- Requires sequential and isolated `WorksheetPart` streaming.
- Low memory constraints must be preserved.

### 2. **Basic Cell Styling API**
Lightweight styling (unlike ClosedXML):
- Bold / italic / underline  
- Background color  
- Borders (simple)  
- Text alignment  
- Header row style options

### 3. **Shared Strings Table (In-Memory Mode)**
- Reduce file size  
- Improve Excel loading speed  
- Cache/lookup strings  
- Use `SharedStringTablePart` for optimal performance  

### 4. **AutoFit 2.0 (Improved Column Width Calculation)**
- Character count → pixel width approximation  
- More accurate widths without COM interop  
- 90–95% match to Excel auto-fit  

---

# 🟧 PHASE 2 — Workbook Functional Enhancements (Medium Priority)

### 5. **Hyperlink Support**
- External URLs  
- Internal sheet references  
- Optional hyperlink-style formatting  

### 6. **Data Validation Support**
- Dropdown lists  
- Number/date ranges  
- Custom formula validation  
- Error messages & user prompts  

### 7. **Worksheet & Workbook Protection**
- Protect structure  
- Lock/unlock specific cells  
- Password hash (SHA‑512) support  

### 8. **Merged Cells**
- Title rows  
- Multi-column headers  
- Grouping UI sections  

---

# 🟨 PHASE 3 — Advanced Excel Features (Long‑Term)

### 9. **Image Embedding (PNG/JPG)**
- Add drawing parts  
- Embedded logos or charts  
- Position images in cells  

### 10. **Basic Chart Support**
Minimal chart engine:
- Line  
- Bar  
- Pie  
- Data series + layout templates  

### 11. **Freeze Panes**
- Freeze top row  
- Freeze first column  
- Freeze custom areas  

### 12. **Conditional Formatting**
- Highlight rules  
- Color scales  
- Icon sets  

---

# 🟥 PHASE 4 — Enterprise‑Level Functionality (Very Long‑Term)

### 13. **Pivot Table Generation**
Requires:
- PivotCache  
- CacheDefinition  
- Field mappings  
- Multi‑part relationships  

This is a complex and heavy feature—only added if demand increases.

---

# 🌀 Developer Experience & Tooling

### 14. **Public API Improvements / Extensions**
- More fluent style  
- Additional helpers  
- Workbook-level configuration  

### 15. **Benchmark & Profiling Suite**
- BenchmarkDotNet  
- Memory pressure tracking  
- Streaming performance metrics  

### 16. **CI/CD Automation**
- GitHub Actions  
- Auto-run OpenXmlValidator  
- Automatic NuGet publish workflow  

### 17. **Documentation Website**
- Full API reference  
- Examples  
- Best practices for huge datasets  

---

# ⭐ Summary Table

| Priority | Feature Group |
|----------|---------------|
| 🔥 High | Basic styles, Shared strings, Multi-sheet streaming, AutoFit 2.0 |
| 🟧 Medium | Hyperlinks, Data validation, Protection, Merged cells |
| 🟨 Long‑Term | Images, Charts, Freeze panes, Conditional formatting |
| 🟥 Very Long-Term | Pivot tables |
| 🌀 DX | CI/CD, Benchmarks, Documentation |

---

# End of Roadmap
