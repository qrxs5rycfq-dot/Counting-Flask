import os
import psycopg2
import json

class BlacklistTracker:
    def __init__(self):
        db_url = os.getenv("DATABASE_URL")
        if not db_url:
            raise ValueError("DATABASE_URL tidak ditemukan.")
        self.conn = psycopg2.connect(db_url)

    def run(self):
        data = {"data": []}
        try:
            with self.conn.cursor() as cur:
                cur.execute("SELECT * FROM acc_person WHERE disabled = 't'")
                acc_persons = cur.fetchall()

                for acc in acc_persons:
                    person_id = acc[17]  # acc_person.person_id

                    cur.execute("SELECT * FROM pers_person WHERE id = %s", (person_id,))
                    person = cur.fetchone()
                    if not person:
                        continue

                    cur.execute("SELECT * FROM pers_attribute_ext WHERE person_id = %s", (person[0],))
                    attribute = cur.fetchone()

                    cur.execute("SELECT * FROM att_person WHERE pers_person_pin = %s", (person[34],))
                    attendance = cur.fetchone()

                    gender = "Male" if person[19] == "M" else "Female"
                    nipeg = attribute[19] if attribute and attribute[19] else ""

                    data["data"].append({
                        "site": "PT. PLN Indonesia Power",
                        "dept": attendance[17] if attendance else "",
                        "foto": "",
                        "name": attendance[24] if attendance else "",
                        "time": gender,
                        "id": attendance[25] if attendance else "",
                        "nipeg": nipeg
                    })
        except Exception as e:
            return {"error": str(e)}

        return data
