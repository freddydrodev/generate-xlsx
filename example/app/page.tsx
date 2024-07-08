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
              { header: "Id", key: "id", width: 10 },
              { header: "Name", key: "name", width: 32 },
              { header: "D.O.B.", key: "dob", width: 10 },
            ],
            data: [
              {
                id: 1,
                name: "John Doe",
                dob: "okok",
              },
              {
                id: 2,
                name: "Jane Doe",
                dob: "pkokl",
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
