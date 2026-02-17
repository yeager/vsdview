# Translations

Translations for VSDView are managed via **Transifex**:

👉 https://www.transifex.com/danielnylander/vsdview/

## How to contribute translations

1. Create a free account on [Transifex](https://www.transifex.com/)
2. Join the [VSDView project](https://www.transifex.com/danielnylander/vsdview/)
3. Translate strings in the web editor
4. Translations are automatically synced to this repository via CI

## ⚠️ Please do NOT submit pull requests for translations

All translation work should be done in Transifex. PRs that modify `.po` files directly will be closed — Transifex is the single source of truth and any manual changes will be overwritten by the next sync.

## For developers

- Source strings are in `vsdview.pot`
- Run `xgettext --from-code=UTF-8 -o po/vsdview.pot vsdview/*.py --keyword=_ --keyword=N_` to regenerate
- The `transifex-sync.yml` workflow handles push/pull automatically
