"use client";

import { generateXLSXGrid } from "@freddydrodev/generate-xlsx";
import styles from "./page.module.css";

export default function Home() {
  return (
    <main className={styles.main}>
      <button
        onClick={async () => {
          await generateXLSXGrid({
            config: {
              name: "My XLSX",
            },
            headers: [
              { header: "ID", key: "id", width: 20 },
              {
                header: "NAME",
                key: "name",
                width: 32,
                alignment: {
                  horizontal: "right",
                  vertical: "middle",
                },
                isCurrency: true,
              },
              {
                header: "ENTREE",
                key: "income",
                width: 30,
                isCurrency: true,
                hasTotal: true,
              },
              {
                header: "SORTIE",
                key: "outcome",
                width: 30,
                isCurrency: true,
                hasTotal: true,
              },
              { header: "DESCRIPTION", key: "description", width: 100 },
            ],
            data: [
              {
                id: 1,
                name: "okok",
                description: "okok",
                income: 20000000,
              },
              {
                id: 2,
                name: "Jane Doe",
                description: "je suis la",
                outcome: -50000000,
              },
            ],
            fileName: "My Fist test",
          });
        }}
      >
        Get XLSX
      </button>
    </main>
  );
}
