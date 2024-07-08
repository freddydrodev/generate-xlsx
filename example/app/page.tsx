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
                hasTotal: true,
              },
              {
                header: "PRICE",
                key: "price",
                width: 30,
                isCurrency: true,
                hasTotal: true,
              },
              { header: "D.O.B.", key: "dob", width: 10 },
            ],
            data: [
              {
                id: 1,
                name: 500,
                dob: "okok",
                price: 2000,
              },
              {
                id: 2,
                name: "Jane Doe",
                dob: 20000,
                price: -500000,
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
