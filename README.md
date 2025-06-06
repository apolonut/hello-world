# hello-world
I am very new in the matter. Can someone help me understand basics in programming? 
#

## Certificate Generation

To generate certificates, ensure that you have a `certificate_template.docx` and a `certificates.csv` with the columns `NAME`, `DATE`, and `CERT_ID` in the repository root. Then install `python-docx` and run:

```bash
pip install python-docx
python generate_certificates.py
```

The generated `.docx` files will appear in the `output/` directory.
