# -*- coding: utf-8 -*-
"""
Constraint Parser - Luonnollisen kielen rajoitteet

Muuntaa käyttäjän ohjeet rajoitteiksi:
- "EU ei voi tehdä yövuoroa" → {"worker": "Dayman EU", "type": "no_night_shift"}
- "PH1 tekee max 8h" → {"worker": "Dayman PH1", "type": "max_hours", "value": 8}
- "Kaikki tarvitsevat 6h lepojakson" → {"type": "min_rest_period", "value": 6}

Käyttää LLM:ää parsimiseen ja validoi tulokset.
"""

import json
import os
from typing import Dict, List, Any, Optional

try:
    import anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False


# Tuetut rajoitetyypit
CONSTRAINT_TYPES = {
    "no_night_shift": {
        "description": "Ei yövuoroa (00:00-08:00)",
        "params": ["worker"],
        "example": "EU ei voi tehdä yövuoroa"
    },
    "no_evening_shift": {
        "description": "Ei iltavuoroa (17:00-00:00)",
        "params": ["worker"],
        "example": "PH1 ei iltavuoroa"
    },
    "max_hours": {
        "description": "Maksimitunnit päivässä",
        "params": ["worker", "value"],
        "example": "PH2 tekee max 8 tuntia"
    },
    "min_hours": {
        "description": "Minimitunnit päivässä",
        "params": ["worker", "value"],
        "example": "EU tekee vähintään 9 tuntia"
    },
    "must_work_slot": {
        "description": "Pakollinen työaika",
        "params": ["worker", "start_time", "end_time"],
        "example": "PH1 on töissä 08:00-12:00"
    },
    "cannot_work_slot": {
        "description": "Kielletty työaika",
        "params": ["worker", "start_time", "end_time"],
        "example": "EU ei voi olla töissä 06:00-08:00"
    },
    "prefer_continuous": {
        "description": "Suosi yhtenäistä vuoroa",
        "params": ["worker"],
        "example": "PH2:lle mieluiten yhtenäinen vuoro"
    },
    "day_off": {
        "description": "Vapaapäivä",
        "params": ["worker", "day"],
        "example": "EU vapaalla päivänä 2"
    },
    "assign_night_shift": {
        "description": "Määrää yövuoroon",
        "params": ["worker", "day"],
        "example": "PH1 tekee yövuoron päivänä 1"
    }
}

# Parser system prompt
PARSER_SYSTEM_PROMPT = """Olet työvuororajoitteiden parseri. Tehtäväsi on muuntaa luonnollisen kielen ohjeet JSON-rajoitteiksi.

TYÖNTEKIJÄT:
- Dayman EU (tai pelkkä EU)
- Dayman PH1 (tai pelkkä PH1)
- Dayman PH2 (tai pelkkä PH2)
- Bosun
- Watchman 1, 2, 3

TUETUT RAJOITETYYPIT:
""" + json.dumps(CONSTRAINT_TYPES, indent=2, ensure_ascii=False) + """

VASTAA AINA VAIN JSON-MUODOSSA:
{
  "constraints": [
    {
      "type": "rajoitetyyppi",
      "worker": "Dayman XX",  // jos tarpeen
      "value": 8,             // jos tarpeen
      "day": 1,               // jos tarpeen
      "start_time": "08:00",  // jos tarpeen
      "end_time": "12:00"     // jos tarpeen
    }
  ],
  "understood": true,
  "clarification_needed": null
}

Jos et ymmärrä ohjetta tai se on epäselvä:
{
  "constraints": [],
  "understood": false,
  "clarification_needed": "Mitä tarkoitit...?"
}

TÄRKEÄÄ:
- Palauta VAIN JSON, ei muuta tekstiä
- Käytä aina täysiä nimiä (Dayman EU, ei pelkkä EU)
- Ajat muodossa HH:MM
- Päivät numeroina (1, 2, 3...)
"""


class ConstraintParser:
    """Parsii luonnollisen kielen rajoitteiksi."""
    
    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or os.environ.get('ANTHROPIC_API_KEY')
        self.client = None
        self.model = "claude-sonnet-4-20250514"
        self.active_constraints = []
        
        if ANTHROPIC_AVAILABLE and self.api_key:
            self.client = anthropic.Anthropic(api_key=self.api_key)
    
    def is_available(self) -> bool:
        return self.client is not None
    
    def parse(self, user_input: str) -> Dict[str, Any]:
        """
        Parsii käyttäjän ohjeen rajoitteiksi.
        
        Args:
            user_input: Käyttäjän ohje luonnollisella kielellä
        
        Returns:
            Dict: {constraints: [...], understood: bool, clarification_needed: str|None}
        """
        if not self.is_available():
            return {
                "constraints": [],
                "understood": False,
                "clarification_needed": "API ei käytettävissä"
            }
        
        try:
            response = self.client.messages.create(
                model=self.model,
                max_tokens=1024,
                system=PARSER_SYSTEM_PROMPT,
                messages=[{"role": "user", "content": user_input}]
            )
            
            response_text = response.content[0].text.strip()
            
            # Poista mahdolliset markdown-kooditagit
            if response_text.startswith("```"):
                lines = response_text.split("\n")
                response_text = "\n".join(lines[1:-1])
            
            result = json.loads(response_text)
            
            # Validoi rajoitteet
            validated = self._validate_constraints(result.get("constraints", []))
            result["constraints"] = validated
            
            return result
            
        except json.JSONDecodeError as e:
            return {
                "constraints": [],
                "understood": False,
                "clarification_needed": f"Parsinta epäonnistui: {str(e)}"
            }
        except Exception as e:
            return {
                "constraints": [],
                "understood": False,
                "clarification_needed": f"Virhe: {str(e)}"
            }
    
    def _validate_constraints(self, constraints: List[Dict]) -> List[Dict]:
        """Validoi ja korjaa rajoitteet."""
        validated = []
        
        for c in constraints:
            if not isinstance(c, dict):
                continue
            
            c_type = c.get("type")
            if c_type not in CONSTRAINT_TYPES:
                continue
            
            # Normalisoi työntekijän nimi
            if "worker" in c:
                c["worker"] = self._normalize_worker_name(c["worker"])
            
            # Normalisoi ajat
            if "start_time" in c:
                c["start_time"] = self._normalize_time(c["start_time"])
            if "end_time" in c:
                c["end_time"] = self._normalize_time(c["end_time"])
            
            validated.append(c)
        
        return validated
    
    def _normalize_worker_name(self, name: str) -> str:
        """Normalisoi työntekijän nimi."""
        name = name.strip()
        
        # Lyhyet nimet täysiksi
        mappings = {
            "EU": "Dayman EU",
            "PH1": "Dayman PH1", 
            "PH2": "Dayman PH2",
            "Dayman EU": "Dayman EU",
            "Dayman PH1": "Dayman PH1",
            "Dayman PH2": "Dayman PH2",
            "Bosun": "Bosun"
        }
        
        return mappings.get(name, name)
    
    def _normalize_time(self, time_str: str) -> str:
        """Normalisoi aika HH:MM muotoon."""
        if not time_str:
            return time_str
        
        time_str = str(time_str).strip()
        
        # Jos pelkkä numero, lisää :00
        if time_str.isdigit():
            return f"{int(time_str):02d}:00"
        
        # Korvaa piste kaksoispisteellä
        time_str = time_str.replace(".", ":")
        
        # Varmista HH:MM muoto
        parts = time_str.split(":")
        if len(parts) == 2:
            h, m = int(parts[0]), int(parts[1])
            return f"{h:02d}:{m:02d}"
        
        return time_str
    
    def add_constraint(self, constraint: Dict) -> bool:
        """Lisää rajoite aktiivisiin."""
        if self._is_valid_constraint(constraint):
            self.active_constraints.append(constraint)
            return True
        return False
    
    def remove_constraint(self, index: int) -> bool:
        """Poista rajoite indeksillä."""
        if 0 <= index < len(self.active_constraints):
            self.active_constraints.pop(index)
            return True
        return False
    
    def clear_constraints(self):
        """Tyhjennä kaikki rajoitteet."""
        self.active_constraints = []
    
    def get_constraints(self) -> List[Dict]:
        """Palauta aktiiviset rajoitteet."""
        return self.active_constraints.copy()
    
    def _is_valid_constraint(self, constraint: Dict) -> bool:
        """Tarkista onko rajoite validi."""
        return (
            isinstance(constraint, dict) and 
            constraint.get("type") in CONSTRAINT_TYPES
        )
    
    def format_constraints(self) -> str:
        """Muotoile rajoitteet luettavaksi."""
        if not self.active_constraints:
            return "Ei aktiivisia rajoitteita."
        
        lines = ["Aktiiviset rajoitteet:"]
        for i, c in enumerate(self.active_constraints):
            desc = self._describe_constraint(c)
            lines.append(f"  {i+1}. {desc}")
        
        return "\n".join(lines)
    
    def _describe_constraint(self, constraint: Dict) -> str:
        """Kuvaile rajoite luonnollisella kielellä."""
        c_type = constraint.get("type")
        worker = constraint.get("worker", "")
        
        descriptions = {
            "no_night_shift": f"{worker}: ei yövuoroa",
            "no_evening_shift": f"{worker}: ei iltavuoroa",
            "max_hours": f"{worker}: max {constraint.get('value')}h",
            "min_hours": f"{worker}: min {constraint.get('value')}h",
            "must_work_slot": f"{worker}: töissä {constraint.get('start_time')}-{constraint.get('end_time')}",
            "cannot_work_slot": f"{worker}: ei töissä {constraint.get('start_time')}-{constraint.get('end_time')}",
            "prefer_continuous": f"{worker}: yhtenäinen vuoro",
            "day_off": f"{worker}: vapaalla päivänä {constraint.get('day')}",
            "assign_night_shift": f"{worker}: yövuoro päivänä {constraint.get('day')}"
        }
        
        return descriptions.get(c_type, str(constraint))
    
    def constraints_to_generator_params(self) -> Dict[str, Any]:
        """
        Muuntaa rajoitteet generaattorin parametreiksi.
        
        Returns:
            Dict joka voidaan välittää generate_schedule() funktiolle
        """
        params = {
            "worker_constraints": {},
            "global_constraints": {}
        }
        
        for c in self.active_constraints:
            worker = c.get("worker")
            c_type = c.get("type")
            
            if worker:
                if worker not in params["worker_constraints"]:
                    params["worker_constraints"][worker] = []
                params["worker_constraints"][worker].append(c)
            else:
                if c_type not in params["global_constraints"]:
                    params["global_constraints"][c_type] = []
                params["global_constraints"][c_type].append(c)
        
        return params


def create_parser(api_key: Optional[str] = None) -> ConstraintParser:
    """Luo uusi parser."""
    return ConstraintParser(api_key)


# ============================================================================
# TESTAUS
# ============================================================================

def test_parser(api_key: Optional[str] = None):
    """Testaa parseria."""
    print("=" * 60)
    print("CONSTRAINT PARSER - TESTI")
    print("=" * 60)
    
    parser = create_parser(api_key)
    
    if not parser.is_available():
        print("\n⚠️ API ei käytettävissä")
        print("Aseta ANTHROPIC_API_KEY ympäristömuuttuja")
        return
    
    print("\nAPI käytettävissä ✓")
    
    # Testilauseet
    test_inputs = [
        "EU ei voi tehdä yövuoroa",
        "PH1 tekee maksimissaan 8 tuntia päivässä",
        "PH2 on töissä kello 10-14",
        "EU on vapaalla päivänä 2",
        "Anna PH1:lle yövuoro päivänä 1"
    ]
    
    for user_input in test_inputs:
        print(f"\n{'='*50}")
        print(f"INPUT: \"{user_input}\"")
        print("-" * 50)
        
        result = parser.parse(user_input)
        
        if result["understood"]:
            print("✓ Ymmärretty")
            for c in result["constraints"]:
                print(f"  → {json.dumps(c, ensure_ascii=False)}")
                parser.add_constraint(c)
        else:
            print(f"✗ Ei ymmärretty: {result['clarification_needed']}")
    
    print(f"\n{'='*60}")
    print("AKTIIVISET RAJOITTEET:")
    print("=" * 60)
    print(parser.format_constraints())
    
    print(f"\n{'='*60}")
    print("GENERAATTORI-PARAMETRIT:")
    print("=" * 60)
    print(json.dumps(parser.constraints_to_generator_params(), indent=2, ensure_ascii=False))


if __name__ == "__main__":
    import sys
    api_key = sys.argv[1] if len(sys.argv) > 1 else None
    test_parser(api_key)
