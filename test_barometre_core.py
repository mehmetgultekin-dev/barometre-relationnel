import unittest

from barometre_core import (
    ProjetInvalide,
    classer_relation,
    longueur_max_colonne,
    valider_donnees_projet,
)


class ValiderDonneesProjetTests(unittest.TestCase):
    def setUp(self):
        self.valid_project = {
            "participants": [
                {"nom": " Alice ", "service": " Équipe A "},
                {"nom": "Bob", "service": "Équipe B"},
            ],
            "services": [],
            "relations_saisies": [{
                "Émetteur": "Alice",
                "Récepteur": "Bob",
                "P+": 1,
            }],
            "nombre_total_personnes": 1,
        }

    def test_normalizes_names_scores_and_total(self):
        result = valider_donnees_projet(self.valid_project)
        relation = result["relations_saisies"][0]
        self.assertEqual(result["participants"][0]["nom"], "Alice")
        self.assertEqual(result["nombre_total_personnes"], 2)
        self.assertEqual(relation["Score Pic Positif"], 1)
        self.assertEqual(relation["Vigilance"], "Positif")

    def test_rejects_non_object_json(self):
        with self.assertRaises(ProjetInvalide):
            valider_donnees_projet([])

    def test_rejects_duplicate_participants(self):
        self.valid_project["participants"].append({"nom": "Alice", "service": "X"})
        with self.assertRaises(ProjetInvalide):
            valider_donnees_projet(self.valid_project)

    def test_rejects_unknown_relation_participant(self):
        self.valid_project["relations_saisies"][0]["Récepteur"] = "Inconnu"
        with self.assertRaises(ProjetInvalide):
            valider_donnees_projet(self.valid_project)

    def test_rejects_invalid_indicator(self):
        self.valid_project["relations_saisies"][0]["P+"] = 2
        with self.assertRaises(ProjetInvalide):
            valider_donnees_projet(self.valid_project)

    def test_rejects_non_numeric_total(self):
        self.valid_project["nombre_total_personnes"] = "deux"
        with self.assertRaises(ProjetInvalide):
            valider_donnees_projet(self.valid_project)

    def test_classification_covers_all_score_combinations(self):
        expected = {
            (0, 0): "Aucune donnée",
            (1, 0): "Positif",
            (2, 0): "Positif",
            (3, 0): "Positif pur",
            (0, 1): "Négatif",
            (0, 2): "Négatif",
            (0, 3): "Négatif pur",
            (1, 1): "Mixte tendu",
            (1, 2): "Mixte tendu",
            (1, 3): "Mixte tendu",
            (2, 1): "Mixte positif",
            (2, 2): "Mixte tendu",
            (2, 3): "Mixte tendu",
            (3, 1): "Mixte positif",
            (3, 2): "Mixte positif",
            (3, 3): "Mixte tendu",
        }
        for scores, vigilance in expected.items():
            with self.subTest(scores=scores):
                self.assertEqual(classer_relation(*scores), vigilance)

    def test_column_width_handles_missing_values(self):
        self.assertEqual(longueur_max_colonne([None, "Unidirectionnelle", ""]), 17)


if __name__ == "__main__":
    unittest.main()