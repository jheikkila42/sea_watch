# Sea Watch

Tämä projekti sisältää Streamlit-sovelluksen (`app.py`) ja työvuorologiikan (`sea_watch_10.py`).

## Miten saat viimeisimmät muutokset Streamlitiin

Streamlit Community Cloud ajaa aina **yhtä valittua haaraa** (teillä `main`).
Siksi muutokset näkyvät vasta, kun ne on mergetty `main`-haaraan.

### Suositeltu workflow (joka kerta)

1. Tee muutokset omaan haaraan (esim. `feature/...` tai `codex/...`).
2. Avaa Pull Request: **base = `main`**, **compare = oma haara**.
3. Varmista että PR:n *Files changed* -näkymässä näkyy oikeasti odottamasi muutokset.
4. Mergeä PR `main`-haaraan.
5. Avaa Streamlitin logit ja varmista että käynnistyksessä lukee:
   - `branch: 'main'`
   - uusi commit SHA / tuore build-aika.
6. Tarvittaessa paina Streamlitissä **Reboot app**.

## Nopea tarkistuslista ennen mergeä

- [ ] PR:ssä `base` on `main`.
- [ ] PR ei sano "There isn’t anything to compare" silloin kun odotat muutoksia.
- [ ] PR:n diffissä näkyy juuri ne tiedostot joita muutit (esim. `sea_watch_10.py`).
- [ ] PR on mergetty `main`-haaraan.
- [ ] Streamlit build-logissa näkyy uusi deploy `main`-haarasta.

## Jos Streamlit ei päivity

Yleisimmät syyt:

- Muutokset on eri haarassa kuin `main`.
- PR on avattu väärään suuntaan (base/compare väärin).
- PR on kyllä tehty, mutta sitä ei ole mergetty.
- Streamlitissä näkyy vanha build (ratkaisu: Reboot app + tarkista logeista branch/commit).


## Kun GitHub sanoo "There isn't anything to compare"

Tämä tarkoittaa yleensä, että valittu `compare`-haara ei sisällä yhtään committia,
jota `main` ei jo sisällä.

Käy läpi tämä järjestyksessä:

1. Avaa PR-sivulla **base = main** ja vaihda **compare**-haaraksi se haara,
   jossa uudet commitit oikeasti ovat.
2. Avaa kyseisen haaran "Commits"-näkymä ja varmista, että siellä näkyy
   tuorein commit (esim. päivämäärä/tunniste).
3. Jos haaraa ei löydy listasta, sitä ei todennäköisesti ole pushattu GitHubiin.
4. Jos PR on jo olemassa mutta "nothing to compare", tee uusi PR oikeasta haarasta.
5. Merge jälkeen tarkista Streamlit-logeista, että build on käynnistynyt uudella
   commitilla `main`-haarasta.

### Käytännön tulkinta nykytilanteeseen

- Streamlit ajaa `main`-haaraa.
- Jos `main` näyttää vanhaa käytöstä, uudet commitit eivät ole vielä oikeasti
  mergetty `main`:iin (tai compare-haara on väärä/identtinen `main`:in kanssa).
