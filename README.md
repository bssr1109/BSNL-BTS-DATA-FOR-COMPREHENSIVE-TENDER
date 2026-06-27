# BTS DATA 

## Use now
1. Supabase > SQL Editor > New query > paste `supabase_schema.sql` > Run.
2. Supabase > Table Editor > `sites` > Import CSV > upload `seed_sites.csv`.
3. In Streamlit project, create `.streamlit/secrets.toml` using `secrets.toml.example`.
4. Run:

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Tables created
- `sites` = BTS master data and field filled data
- `dropdown_options` = dropdown choices used by Streamlit app

## Field login
- Select SDCA
- Enter Name of SDE CM/CFA
- Fill required fields
- Save All button enables only after all required fields are filled

## Master login
- Password protected
- Combined report
- SDCA completion summary
- Excel download

Rows included from uploaded Excel: 98
