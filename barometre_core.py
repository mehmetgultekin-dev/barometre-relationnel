"""Validation et calculs métier indépendants de Streamlit."""


INDICATORS = ("P+", "P-", "I+", "I-", "C+", "C-")


class ProjetInvalide(ValueError):
    """Raised when imported project data cannot be used safely."""


def classer_relation(p_plus: int, p_moins: int) -> str:
    if p_plus == 3 and p_moins == 0:
        return "Positif pur"
    if p_plus == 0 and p_moins == 3:
        return "Négatif pur"
    if p_plus in (1, 2) and p_moins == 0:
        return "Positif"
    if p_plus == 0 and p_moins in (1, 2):
        return "Négatif"
    if p_plus > 0 and p_moins > 0 and p_plus > p_moins:
        return "Mixte positif"
    if p_plus > 0 and p_moins > 0:
        return "Mixte tendu"
    return "Aucune donnée"


def valider_donnees_projet(data):
    if not isinstance(data, dict):
        raise ProjetInvalide("Le fichier doit contenir un objet projet JSON.")

    participants_source = data.get("participants", [])
    relations_source = data.get("relations_saisies", [])
    services_source = data.get("services", [])
    if not isinstance(participants_source, list) or not isinstance(relations_source, list):
        raise ProjetInvalide("Les participants et relations doivent être des listes.")
    if not isinstance(services_source, list):
        raise ProjetInvalide("La liste des services est invalide.")

    participants = []
    noms = set()
    for participant in participants_source:
        if not isinstance(participant, dict):
            raise ProjetInvalide("Un participant du fichier est invalide.")
        nom = participant.get("nom")
        service = participant.get("service")
        if not isinstance(nom, str) or not nom.strip():
            raise ProjetInvalide("Chaque participant doit avoir un nom valide.")
        if not isinstance(service, str) or not service.strip():
            raise ProjetInvalide("Chaque participant doit avoir un service valide.")
        nom = nom.strip()
        service = service.strip()
        if nom in noms:
            raise ProjetInvalide(f"Le participant « {nom} » apparaît plusieurs fois.")
        noms.add(nom)
        participants.append({"nom": nom, "service": service})

    relations = []
    identifiants = set()
    for relation in relations_source:
        if not isinstance(relation, dict):
            raise ProjetInvalide("Une relation du fichier est invalide.")
        emetteur = relation.get("Émetteur")
        recepteur = relation.get("Récepteur")
        if not isinstance(emetteur, str) or not isinstance(recepteur, str):
            raise ProjetInvalide("Une relation doit préciser son émetteur et son récepteur.")
        emetteur = emetteur.strip()
        recepteur = recepteur.strip()
        if emetteur not in noms or recepteur not in noms or emetteur == recepteur:
            raise ProjetInvalide("Une relation référence un participant inconnu ou identique.")

        relation_normalisee = {
            "Émetteur": emetteur,
            "Récepteur": recepteur,
        }
        for indicator in INDICATORS:
            value = relation.get(indicator, 0)
            if isinstance(value, bool):
                value = int(value)
            if not isinstance(value, int) or value not in (0, 1):
                raise ProjetInvalide(f"L’indicateur {indicator} doit être égal à 0 ou 1.")
            relation_normalisee[indicator] = value

        p_plus = sum(relation_normalisee[key] for key in ("P+", "I+", "C+"))
        p_moins = sum(relation_normalisee[key] for key in ("P-", "I-", "C-"))
        relation_normalisee.update({
            "Date": str(relation.get("Date", "RAS")),
            "Début": str(relation.get("Début", "RAS")),
            "Fin": str(relation.get("Fin", "RAS")),
            "Service": str(relation.get("Service", "RAS")),
            "Score Pic Positif": p_plus,
            "Score Pic Négatif": p_moins,
            "Score Net": p_plus - p_moins,
            "Vigilance": classer_relation(p_plus, p_moins),
            "Commentaire": str(relation.get("Commentaire", "")),
        })

        identifier = tuple(relation_normalisee[key] for key in (
            "Émetteur", "Récepteur", "Date", "Début", "Fin"
        ))
        if identifier in identifiants:
            raise ProjetInvalide("Le fichier contient une relation enregistrée en double.")
        identifiants.add(identifier)
        relations.append(relation_normalisee)

    total = data.get("nombre_total_personnes", len(participants))
    if isinstance(total, bool) or not isinstance(total, int) or total < 0:
        raise ProjetInvalide("Le nombre total de personnes doit être un entier positif ou nul.")
    total = max(total, len(participants))

    services = []
    for service in services_source:
        if not isinstance(service, str):
            raise ProjetInvalide("Un nom de service est invalide.")
        if service.strip() and service.strip() not in services:
            services.append(service.strip())
    for participant in participants:
        if participant["service"] not in services:
            services.append(participant["service"])

    return {
        "participants": participants,
        "services": services,
        "relations_saisies": relations,
        "nombre_total_personnes": total,
    }