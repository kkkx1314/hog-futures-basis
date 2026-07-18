
import time, pandas as pd
from pathlib import Path
import akshare as ak

BASE = Path('data')
FUTURES_DIR = BASE / 'futures'
HOLDINGS_DIR = BASE / 'holdings'
HOLDINGS_DIR.mkdir(exist_ok=True)

contracts = ['LH2501', 'LH2505', 'LH2507', 'LH2509', 'LH2511', 'LH2601', 'LH2605', 'LH2607', 'LH2609', 'LH2611', 'LH2701', 'LH2705']
CUTOFF = '20250410'

for ct in contracts:
    df = pd.read_csv(FUTURES_DIR / f'{ct}.csv')
    df["date"] = pd.to_datetime(df["date"])
    dates = sorted(df[df["date"] >= CUTOFF]["date"].dt.strftime('%Y%m%d').unique())
    missing = [d for d in dates if not (HOLDINGS_DIR / f'{ct}_{d}.csv').exists()]
    if not missing:
        print(f'{ct}: all {len(dates)} cached')
        continue
    print(f'{ct}: {len(dates)} dates, {len(missing)} missing')
    for i, d in enumerate(missing):
        cache = HOLDINGS_DIR / f'{ct}_{d}.csv'
        if cache.exists():
            continue
        try:
            dv = ak.futures_hold_pos_sina(symbol='成交量', contract=ct, date=d)
            dl = ak.futures_hold_pos_sina(symbol='多单持仓', contract=ct, date=d)
            ds = ak.futures_hold_pos_sina(symbol='空单持仓', contract=ct, date=d)
            # Quick check
            if len(dl) == 0:
                continue
            def _n(df, vc, cc):
                out = pd.DataFrame()
                for c in df.columns:
                    if '会员' in str(c) or '简称' in str(c):
                        out['company'] = df[c].astype(str).str.strip()
                        break
                for c in df.columns:
                    if '名次' not in str(c) and '会员' not in str(c) and '简称' not in str(c) and '增减' not in str(c) and '比上' not in str(c):
                        out[vc] = pd.to_numeric(df[c], errors='coerce').fillna(0).astype(int)
                        break
                for c in df.columns:
                    if '增减' in str(c) or '比上' in str(c):
                        out[cc] = pd.to_numeric(df[c], errors='coerce').fillna(0).astype(int)
                        break
                if cc not in out.columns: out[cc] = 0
                return out[out['company'].notna() & (out['company'] != '')]
            vv = _n(dv, 'volume', 'volume_chg')
            ll = _n(dl, 'long', 'long_chg')
            ss = _n(ds, 'short', 'short_chg')
            m = ll.merge(ss, on='company', how='outer').merge(vv, on='company', how='outer').fillna(0)
            if m['long'].sum() == 0:
                continue
            m.to_csv(cache, index=False)
            gen = HOLDINGS_DIR / f'{ct}.csv'
            m.to_csv(gen, index=False)
            (HOLDINGS_DIR / f'{ct}_meta.txt').write_text(d)
        except:
            pass
        if (i+1) % 20 == 0:
            print(f'  {ct}: {i+1}/{len(missing)}')
        if i < len(missing) - 1:
            time.sleep(0.1)
    print(f'{ct}: done')
print('All done!')
