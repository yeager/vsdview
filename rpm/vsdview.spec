Name:           vsdview
Version:        0.5.0
Release:        1%{?dist}
Summary:        Read-only viewer for Microsoft Visio files
License:        GPL-3.0-or-later
URL:            https://github.com/yeager/vsdview
Source0:        %{name}-%{version}.tar.gz
BuildArch:      noarch

Requires:       python3 >= 3.10
Requires:       python3-gobject
Requires:       gtk4
Requires:       libadwaita
Requires:       librsvg2
Requires:       libvisio-ng >= 0.6.0
Recommends:     librsvg2-tools

%description
VSDView is a read-only viewer for Microsoft Visio files (.vsdx and .vsd),
built with GTK4 and libadwaita. Features include multi-page viewing,
interactive zoom, pan/drag, text search, shape info panel, measurement tool,
minimap, export to PNG/PDF/SVG, and keyboard shortcuts.

%prep
%autosetup

%install
mkdir -p %{buildroot}%{python3_sitelib}/%{name}
mkdir -p %{buildroot}%{_bindir}
mkdir -p %{buildroot}%{_datadir}/applications
mkdir -p %{buildroot}%{_datadir}/icons/hicolor/scalable/apps
mkdir -p %{buildroot}%{_mandir}/man1

cp vsdview/*.py %{buildroot}%{python3_sitelib}/%{name}/
printf '#!/usr/bin/env python3\nfrom vsdview.__main__ import main\nmain()\n' > %{buildroot}%{_bindir}/vsdview
chmod 755 %{buildroot}%{_bindir}/vsdview

install -m 644 data/se.danielnylander.vsdview.desktop %{buildroot}%{_datadir}/applications/
install -m 644 data/icons/se.danielnylander.vsdview.svg %{buildroot}%{_datadir}/icons/hicolor/scalable/apps/ 2>/dev/null || true
install -m 644 man/vsdview.1 %{buildroot}%{_mandir}/man1/ 2>/dev/null || true

%files
%license LICENSE
%doc README.md CHANGELOG.md
%{_bindir}/vsdview
%{python3_sitelib}/%{name}/
%{_datadir}/applications/se.danielnylander.vsdview.desktop
%{_datadir}/icons/hicolor/scalable/apps/se.danielnylander.vsdview.svg
%{_mandir}/man1/vsdview.1*

%changelog
* Sat Mar 01 2026 Daniel Nylander <daniel@danielnylander.se> - 0.5.0-1
- Major feature release: interactive zoom, pan, shape info, measurement, minimap, SVG export
