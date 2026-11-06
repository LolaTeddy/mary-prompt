# Mary Prompt – Einrichtung

Mary Prompt ist eine Web-App (PWA). Sie muss im Internet über **https** erreichbar sein,
sonst erlaubt das iPhone weder Kamera noch Mikrofon. Am einfachsten und kostenlos geht das mit **GitHub Pages**.

## 1. App online stellen (GitHub Pages)

1. Auf github.com anmelden und ein neues Repository anlegen, z. B. `mary-prompt` (Sichtbarkeit: Public).
2. Im Repository auf „Add file“ › „Upload files“ klicken und **den Inhalt** dieses Ordners hochladen
   (index.html, app.js, styles.css, onedrive.js, speech.js, config.js, sw.js, manifest.webmanifest und den Ordner icons).
3. Settings › Pages › „Deploy from a branch“, Branch `main`, Ordner `/ (root)`, speichern.
4. Nach ein bis zwei Minuten ist die App erreichbar, z. B. unter
   `https://DEIN-NAME.github.io/mary-prompt/`

## Skripte ohne Microsoft-Registrierung (Standard)

„Skript aus Dateien öffnen“ öffnet die Dateien-App des iPhones. Dort OneDrive (OneDrive-App muss installiert
und in der Dateien-App eingeschaltet sein) oder iCloud Drive wählen, dann die Word-Datei (`.docx`) oder Textdatei (`.txt`) antippen.
Alte Word-Dateien (`.doc`) gehen nicht; Word-Dateien brauchen iOS 16.4 oder neuer.
Das Skript wird in der App gespeichert. Wählst du eine Datei mit gleichem Namen erneut, wird die alte Fassung ersetzt.
Die OneDrive-Anbindung unten ist nur nötig, wenn neue Skripte automatisch in der Liste erscheinen sollen.
Sie braucht ein Geschäftskonto (Microsoft 365) oder ein kostenloses Azure-Konto.

## 2. OneDrive anbinden (optional, kostenlos)

1. https://entra.microsoft.com öffnen (mit dem Microsoft-Konto, zu dem dein OneDrive gehört)
   › „App-Registrierungen“ › „Neue Registrierung“.
2. Name: `Mary Prompt`.
3. Unterstützte Kontotypen: **„Konten in einem beliebigen Organisationsverzeichnis und persönliche Microsoft-Konten“**.
4. Umleitungs-URI: Plattform **„Single-Page-Anwendung (SPA)“**, Adresse genau wie die App-Adresse,
   z. B. `https://DEIN-NAME.github.io/mary-prompt/` (mit Schrägstrich am Ende).
   Die exakte Adresse zeigt die App auch unter Einstellungen › OneDrive an.
5. „Registrieren“. Danach die **Anwendungs-ID (Client-ID)** kopieren.
6. Unter „API-Berechtigungen“ sollte `User.Read` stehen. `Files.Read` wird bei der ersten Anmeldung abgefragt,
   du musst nichts weiter eintragen.
7. Die Client-ID entweder in `config.js` bei `clientId` eintragen (dann gilt sie für alle Geräte)
   oder in der App unter Einstellungen › OneDrive einfügen.

In OneDrive einen Ordner **„Teleprompter“** anlegen (direkt in „Eigene Dateien“). Dort die Skripte
als `.txt`-Dateien ablegen. Ein anderer Ordnername lässt sich in den Einstellungen festlegen.

## 3. Auf dem iPhone installieren

1. Die App-Adresse in **Safari** öffnen.
2. Teilen-Knopf › „Zum Home-Bildschirm“.
3. Mary Prompt vom Home-Bildschirm aus starten, „Mit OneDrive verbinden“ tippen und anmelden.
4. Beim ersten Öffnen eines Skripts Kamera und Mikrofon erlauben.

Wichtig: Die Home-Bildschirm-App hat eigene Daten. Die Anmeldung also **in der installierten App** machen,
nicht nur in Safari.

## So funktioniert es

- **Skripte**: Die Liste lädt sich beim Öffnen der App automatisch neu. Neue Dateien im OneDrive-Ordner
  erscheinen ohne Abtippen. Zusätzlich gibt es „Text einfügen“ zum Kopieren und Einfügen.
- **Aufnahme**: Skript antippen › oranger Knopf › Countdown (5–10 Sekunden, einstellbar) › Aufnahme läuft.
  Tippen aufs Bild hält den Text an und lässt ihn weiterlaufen. Am Text ziehen verschiebt ihn.
  „A−/A+“ ändert die Schrift, „−/+“ das Tempo, auch während der Aufnahme.
- **Speichern**: Nach dem Stopp kommt die Vorschau. „Speichern“ öffnet das Teilen-Menü, dort
  **„Video sichern“** wählen, dann liegt das Video in der Fotos-App. „Verwerfen“ löscht die Aufnahme.
- Der Teleprompter-Text ist **nicht** im Video. Das Video wird immer im Hochformat 9:16 aufgenommen.
  Die Vorschau ist gespiegelt (wie ein Spiegel), das gespeicherte Video ist seitenrichtig.

## Hinweise und Grenzen

- **OneDrive-Anmeldung**: Microsoft lässt Web-Apps nur 24 Stunden ohne neue Anmeldung zugreifen.
  Danach erscheint „Anmeldung abgelaufen“. Ein Tipp auf „Mit OneDrive verbinden“ genügt meist,
  ein Passwort ist nur selten nötig. Bereits geladene Skripte bleiben auch offline verfügbar.
- **Sprachsteuerung (Test)**: In den Einstellungen unter „Textlauf“ einschalten. Auf dem iPhone muss
  unter Einstellungen › Allgemein › Tastatur „Diktieren“ aktiv sein. Safari erlaubt die Spracherkennung
  nur eingeschränkt, und das Mikrofon wird gleichzeitig für das Video gebraucht. Ob beides zusammen
  zuverlässig klappt, muss am iPhone getestet werden. Klappt es nicht, schaltet die App automatisch
  auf den Lauf mit eingestelltem Tempo um. Zum Ausprobieren gibt es „Spracherkennung testen“.
- **Bildschirm bleibt an** während der Aufnahme (ab iOS 16.4).
- **Updates**: Nach Änderungen an der App in `sw.js` die Versionsnummer bei `CACHE` erhöhen
  und die Dateien neu hochladen.
