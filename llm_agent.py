# -*- coding: utf-8 -*-
"""
LLM Agent - Claude API integraatio työvuorojen analysointiin

Käyttää Claude API:a:
1. Analysoimaan työvuoro-ongelmia
2. Generoimaan korjausehdotuksia
3. Vastaamaan käyttäjän kysymyksiin
"""

import json
import os
from typing import Dict, List, Any, Optional

try:
    import anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False
    print("Varoitus: anthropic-kirjasto ei asennettu. Asenna: pip install anthropic")


# System prompt joka kuvaa agentin roolin
SYSTEM_PROMPT = """Olet Sea Watch -työvuorosuunnittelun asiantuntija-avustaja. 

Tehtäväsi on analysoida laivan miehistön työvuoroja ja antaa korjausehdotuksia.

TAUSTATIETO:
- Daymanit (EU, PH1, PH2) tekevät päivätyötä satamassa
- Jokaisen daymanin pitää tehdä 8-10h päivässä
- STCW-säännöt: vähintään 10h lepoa 24h aikana, jaettuna max 2 jaksoon joista pisin >= 6h
- Normaalityöaika: 08:00-17:00 (lounas 11:30-12:00)
- Satamaoperaatioiden aikana vähintään 1 dayman töissä
- Tulo/lähtö: kaikki daymanit 1h
- Slussi: 2h (1. tunti 2 dm, 2. tunti 3 dm)
- Shiftaus: 1h kaikki daymanit

OHJEITA:
- Vastaa AINA suomeksi
- Ole ytimekäs ja käytännöllinen
- Anna konkreettisia korjausehdotuksia (kellonajat, työntekijät)
- Perustele ehdotukset STCW-sääntöjen tai käytännön kannalta
- Jos ongelmia ei ole, sano se lyhyesti
"""


class LLMAgent:
    """Claude API -pohjainen agentti työvuoroanalyysiin."""
    
    def __init__(self, api_key: Optional[str] = None):
        """
        Alustaa agentin.
        
        Args:
            api_key: Anthropic API-avain. Jos None, yrittää lukea ANTHROPIC_API_KEY ympäristömuuttujasta.
        """
        self.api_key = api_key or os.environ.get('ANTHROPIC_API_KEY')
        self.client = None
        self.model = "claude-sonnet-4-20250514"
        self.conversation_history = []
        
        if ANTHROPIC_AVAILABLE and self.api_key:
            self.client = anthropic.Anthropic(api_key=self.api_key)
    
    def is_available(self) -> bool:
        """Tarkistaa onko LLM käytettävissä."""
        return self.client is not None
    
    def _call_api(self, messages: List[Dict], max_tokens: int = 1024) -> str:
        """
        Kutsuu Claude API:a.
        
        Returns:
            Vastaus tekstinä tai virheilmoitus.
        """
        if not self.is_available():
            return "VIRHE: Claude API ei käytettävissä. Tarkista API-avain."
        
        try:
            response = self.client.messages.create(
                model=self.model,
                max_tokens=max_tokens,
                system=SYSTEM_PROMPT,
                messages=messages
            )
            return response.content[0].text
        except Exception as e:
            return f"VIRHE API-kutsussa: {str(e)}"
    
    def analyze_and_suggest(self, analysis_data: Dict) -> str:
        """
        Analysoi työvuorot ja antaa korjausehdotuksia.
        
        Args:
            analysis_data: get_analysis_for_llm() palauttama data
        
        Returns:
            Korjausehdotukset tekstinä
        """
        if not analysis_data.get('has_problems', False):
            return "✅ Työvuoroissa ei havaittu ongelmia. Kaikki näyttää hyvältä!"
        
        # Rakenna prompt
        prompt = f"""Analysoi seuraavat työvuoro-ongelmat ja anna korjausehdotukset:

ONGELMAT:
{json.dumps(analysis_data['problems'], indent=2, ensure_ascii=False)}

YHTEENVETO:
- Päiviä: {analysis_data['num_days']}
- STCW-rikkeitä: {analysis_data['summary']['stcw_violations']}
- Op-kattavuusaukkoja: {analysis_data['summary']['op_coverage_gaps']}
- Tuntitasapaino-ongelmia: {analysis_data['summary']['hour_imbalances']}

Anna konkreettiset korjausehdotukset. Kerro:
1. Mikä on ongelma
2. Miten se korjataan (kellonajat, kuka tekee mitä)
3. Miksi tämä korjaus toimii
"""
        
        messages = [{"role": "user", "content": prompt}]
        
        return self._call_api(messages)
    
    def answer_question(self, question: str, schedule_context: Dict = None) -> str:
        """
        Vastaa käyttäjän kysymykseen työvuoroista.
        
        Args:
            question: Käyttäjän kysymys
            schedule_context: Valinnainen konteksti nykyisistä vuoroista
        
        Returns:
            Vastaus tekstinä
        """
        # Rakenna konteksti
        context = ""
        if schedule_context:
            context = f"""
NYKYINEN TILANNE:
{json.dumps(schedule_context, indent=2, ensure_ascii=False)}

"""
        
        prompt = f"""{context}KÄYTTÄJÄN KYSYMYS:
{question}

Vastaa kysymykseen perustuen työvuorosuunnittelun sääntöihin ja käytäntöihin."""
        
        # Lisää keskusteluhistoriaan
        self.conversation_history.append({"role": "user", "content": prompt})
        
        # Kutsu API
        response = self._call_api(self.conversation_history)
        
        # Tallenna vastaus historiaan
        self.conversation_history.append({"role": "assistant", "content": response})
        
        return response
    
    def get_schedule_summary(self, all_days: Dict, days_data: List[Dict]) -> str:
        """
        Generoi yhteenveto työvuoroista luonnollisella kielellä.
        
        Args:
            all_days: Generaattorin palauttama data
            days_data: Alkuperäinen syöte
        
        Returns:
            Yhteenveto tekstinä
        """
        # Kerää perustiedot
        num_days = len(days_data)
        daymen = ['Dayman EU', 'Dayman PH1', 'Dayman PH2']
        
        schedule_info = []
        for d in range(num_days):
            day_info = {'day': d + 1, 'workers': {}}
            
            info = days_data[d]
            arr = info.get('arrival_hour')
            dep = info.get('departure_hour')
            op_start = info.get('port_op_start_hour')
            op_end = info.get('port_op_end_hour')
            
            day_info['arrival'] = f"{arr:02d}:00" if arr else None
            day_info['departure'] = f"{dep:02d}:00" if dep else None
            day_info['op_start'] = f"{op_start:02d}:00" if op_start else None
            day_info['op_end'] = f"{op_end:02d}:00" if op_end is not None else None
            
            for dm in daymen:
                work = all_days[dm][d]['work_slots']
                hours = sum(work) / 2
                
                # Laske työjaksot
                ranges = []
                start = None
                for i, w in enumerate(work):
                    if w and start is None:
                        start = i
                    elif not w and start is not None:
                        h1, m1 = start // 2, "30" if start % 2 else "00"
                        h2, m2 = i // 2, "30" if i % 2 else "00"
                        ranges.append(f"{h1:02d}:{m1}-{h2:02d}:{m2}")
                        start = None
                if start is not None:
                    h1, m1 = start // 2, "30" if start % 2 else "00"
                    ranges.append(f"{h1:02d}:{m1}-00:00")
                
                day_info['workers'][dm] = {
                    'hours': hours,
                    'ranges': ranges
                }
            
            schedule_info.append(day_info)
        
        prompt = f"""Tee lyhyt yhteenveto seuraavasta työvuorolistasta:

{json.dumps(schedule_info, indent=2, ensure_ascii=False)}

Kerro:
1. Montako päivää, mitkä tapahtumat (tulot/lähdöt)
2. Miten työt on jaettu daymanien kesken
3. Onko jotain huomioitavaa (yövuorot, pitkät päivät tms.)

Pidä vastaus lyhyenä (max 5-6 riviä)."""
        
        messages = [{"role": "user", "content": prompt}]
        
        return self._call_api(messages)
    
    def clear_history(self):
        """Tyhjentää keskusteluhistorian."""
        self.conversation_history = []


def create_agent(api_key: Optional[str] = None) -> LLMAgent:
    """
    Luo uuden LLM-agentin.
    
    Args:
        api_key: API-avain (valinnainen, voi lukea ympäristömuuttujasta)
    
    Returns:
        LLMAgent-instanssi
    """
    return LLMAgent(api_key)


# ============================================================================
# TESTAUS (ilman API-kutsua)
# ============================================================================

def test_without_api():
    """Testaa moduulin toimivuus ilman API-kutsua."""
    print("=" * 60)
    print("LLM Agent - Testaus (ilman API-kutsua)")
    print("=" * 60)
    
    agent = create_agent()
    
    print(f"\nAPI käytettävissä: {agent.is_available()}")
    print(f"Malli: {agent.model}")
    
    # Testaa analyysidata
    test_analysis = {
        'num_days': 2,
        'summary': {
            'total_issues': 2,
            'stcw_violations': 0,
            'op_coverage_gaps': 0,
            'hour_imbalances': 1
        },
        'problems': [
            {
                'type': 'worker_issue',
                'worker': 'Dayman PH1',
                'day': 2,
                'hours': 3.0,
                'issues': ['Liian vähän tunteja: 3.0h (min 8h)'],
                'work_ranges': ['00:00-01:00', '06:00-08:00']
            }
        ],
        'has_problems': True
    }
    
    print("\nTestidata:")
    print(json.dumps(test_analysis, indent=2, ensure_ascii=False))
    
    if agent.is_available():
        print("\nKutsutaan API:a...")
        response = agent.analyze_and_suggest(test_analysis)
        print("\nVastaus:")
        print(response)
    else:
        print("\nAPI ei käytettävissä - ohitetaan API-kutsu")
        print("Aseta ANTHROPIC_API_KEY ympäristömuuttuja testataksesi API:a")


if __name__ == "__main__":
    test_without_api()
