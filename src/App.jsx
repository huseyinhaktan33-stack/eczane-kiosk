import { useState, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";

// ── Supabase bağlantısı ──
// URL ve "publishable" anahtar istemci tarafında görünmesi güvenli olan
// herkese açık anahtarlardır (gizli service_role anahtarı DEĞİLDİR).
const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL || "https://wtbhyeshiqufnjagxvxu.supabase.co";
const SUPABASE_KEY = import.meta.env.VITE_SUPABASE_KEY || "sb_publishable_SMPQ0erqqRNEoboVnlvIew_yxsA2uRj";
const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

// ══════════════════════════════════════════════════════════════
// SABİTLER
// ══════════════════════════════════════════════════════════════
const DEFAULT_SUPERADMIN = { email:"admin@eczane.com", password:"Admin123!" };
const MIN_MARGIN = 0.20;
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const genTempPassword = () => {
  const chars = "ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789";
  let out = "";
  for (let i = 0; i < 10; i++) out += chars[Math.floor(Math.random() * chars.length)];
  return out;
};

const BRAND_THEMES = {
  "Avène":          {primary:"#1A5276",accent:"#2E86C1",bg:"#EBF5FB",surface:"#FFF",name:"Avène Mavisi"},
  "Caudalie":       {primary:"#6B2D5E",accent:"#A9527C",bg:"#FAF0F6",surface:"#FFF",name:"Caudalie Mürveri"},
  "La Roche-Posay": {primary:"#4A235A",accent:"#7D3C98",bg:"#F5EEF8",surface:"#FFF",name:"LRP Mor"},
  "Bioderma":       {primary:"#784212",accent:"#CA6F1E",bg:"#FEF9E7",surface:"#FFF",name:"Bioderma Amber"},
  "Vichy":          {primary:"#0E6655",accent:"#17A589",bg:"#E8F8F5",surface:"#FFF",name:"Vichy Turkuaz"},
  "Ducray":         {primary:"#1E8449",accent:"#27AE60",bg:"#EAFAF1",surface:"#FFF",name:"Ducray Yeşil"},
  "CeraVe":         {primary:"#1B2631",accent:"#2E4057",bg:"#EBF5FB",surface:"#FFF",name:"CeraVe Lacivert"},
  "ISDIN":          {primary:"#7D6608",accent:"#B7950B",bg:"#FEFAEE",surface:"#FFF",name:"ISDIN Altın"},
  "La Rosée":       {primary:"#5D4037",accent:"#8D6E63",bg:"#FBF8F5",surface:"#FFF",name:"La Rosée Toprak"},
  "The Purest":     {primary:"#1A237E",accent:"#3949AB",bg:"#EEF0FB",surface:"#FFF",name:"The Purest İndigo"},
  "Dermoskin":      {primary:"#1B5E20",accent:"#388E3C",bg:"#F1F8E9",surface:"#FFF",name:"Dermoskin Zümrüt"},
  "Nötr":           {primary:"#2C4A3E",accent:"#C4974A",bg:"#F7F5F1",surface:"#FFF",name:"Nötr"},
};
const DT = BRAND_THEMES["Nötr"];

const CERT_META = {
  "dermo-test":     {label:"Dermo-test",   color:"#1B4F72",bg:"#D6EAF8"},
  "hypoallergenic": {label:"Hipoalerjenik",color:"#2952A3",bg:"#E8F0FF"},
  "vegan":          {label:"Vegan",        color:"#4A1F00",bg:"#FDEBD0"},
  "spf":            {label:"SPF Korumalı", color:"#7D4700",bg:"#FEF3E0"},
  "fragrance-free": {label:"Parfümsüz",    color:"#4A3D6B",bg:"#F0EBF8"},
  "gmp":            {label:"GMP",          color:"#1E6B3E",bg:"#E9F7EF"},
  "natural":        {label:"Doğal",        color:"#5D4037",bg:"#F5EDE8"},
};

const BOYCOTT_META = {
  temiz:      {label:"✓ Boykot Yok",bg:"#E9F7EF",color:"#1E6B3E"},
  boykot:     {label:"⛔ Boykot",   bg:"#FDECEA",color:"#8B2E2E"},
  bilinmiyor: {label:"? Bilinmiyor",bg:"#F4F4F4",color:"#666"},
};

const PREG_META = {
  safe:    {label:"✓ Hamile Uyumlu",  bg:"#E9F7EF",color:"#1E6B3E"},
  avoid:   {label:"⛔ Hamilede Kaçın",bg:"#FDECEA",color:"#8B2E2E"},
  consult: {label:"⚠ Doktora Danış", bg:"#FFF3E0",color:"#7D4700"},
  unknown: {label:"? Belirtilmemiş",  bg:"#F4F4F4",color:"#666"},
};

const CAT_LABELS = {
  temizleyici:"Temizleyici",serum:"Serum",nemlendirici:"Nemlendirici",
  tonik:"Tonik",spf:"Güneş Kremi",tedavi:"Tedavi",goz:"Göz Kremi",maske:"Maske",
};

const SKIN_TYPE_LABELS = {oily:"Yağlı-Karma",dry:"Kuru",sensitive:"Hassas",normal:"Normal"};
const CONCERN_LABELS = {moisture:"Nem",acne:"Akne & Gözenek",pigment:"Leke & Ton",antiage:"Anti-age",sun:"Güneş Koruması"};

// ══════════════════════════════════════════════════════════════
// İÇERİK (AKTİF MADDE) BİLGİ TABANI — algoritmik cilt tipi / endişe /
// hamilelik uygunluğu önerisi için. Kural tabanlı bir yardımcıdır,
// kesin bir dermatolojik/farmasötik değerlendirme değildir; öneriler
// eczacı/hekim onayıyla teyit edilmelidir. Belirsiz durumlarda
// temkinli taraf seçilir (safe yerine consult).
// ══════════════════════════════════════════════════════════════
const INGREDIENT_KB = [
  {match:["hyalüronik","hyaluronic"], skin_types:["dry","normal","sensitive","oily"], concerns:["moisture"], pregnancy:"safe", label:"Hyalüronik Asit"},
  {match:["niasinamid","niacinamide","vitamin pp"], skin_types:["oily","normal"], concerns:["acne","pigment"], pregnancy:"safe", label:"Niasinamid"},
  {match:["retinol","retinal","retinaldehid","tretinoin","adapalen","vitamin a"], skin_types:["normal","oily"], concerns:["antiage","acne"], pregnancy:"avoid", label:"Retinoid (Retinol/Retinal/Vit. A)"},
  {match:["salisilik","salicylic"," lha","lha "], skin_types:["oily"], concerns:["acne"], pregnancy:"consult", label:"Salisilik Asit (BHA)"},
  {match:["vitamin c","askorbik"], skin_types:["normal","dry","oily"], concerns:["pigment","antiage"], pregnancy:"safe", label:"C Vitamini"},
  {match:["seramid","ceramid"], skin_types:["dry","sensitive"], concerns:["moisture"], pregnancy:"safe", label:"Seramid"},
  {match:["pantenol","panthenol","b5 vitamini","vitamin b5"], skin_types:["sensitive","dry"], concerns:["moisture"], pregnancy:"safe", label:"Pantenol (B5)"},
  {match:["centella","madecassoside","bisabolol","calendula","i-modulia"], skin_types:["sensitive"], concerns:["moisture"], pregnancy:"safe", label:"Sakinleştirici bitkisel özler"},
  {match:["squalane","skualan"], skin_types:["dry"], concerns:["moisture"], pregnancy:"safe", label:"Squalane"},
  {match:["glikolik","laktik asit","mandelik"], skin_types:["normal","oily"], concerns:["antiage","pigment"], pregnancy:"consult", label:"AHA (Glikolik/Laktik Asit)"},
  {match:["benzoil peroksit","benzoyl peroxide"], skin_types:["oily"], concerns:["acne"], pregnancy:"consult", label:"Benzoil Peroksit"},
  {match:["traneksamik","tranexamic"], skin_types:["normal"], concerns:["pigment"], pregnancy:"consult", label:"Traneksamik Asit"},
  {match:["zinc oxide","titanium dioxide","çinko oksit","titanyum dioksit"], skin_types:["sensitive"], concerns:["sun"], pregnancy:"safe", label:"Mineral SPF (Zinc/Titanium Oxide)"},
  {match:["peptid","kolajen","collagen"], skin_types:["normal"], concerns:["antiage"], pregnancy:"safe", label:"Peptid / Kolajen"},
  {match:["ferulik","resveratrol","polifenol"], skin_types:["normal"], concerns:["antiage"], pregnancy:"safe", label:"Antioksidanlar (Ferulik/Resveratrol)"},
  {match:["aloe vera","gül suyu","salatalık","üzüm suyu","grape water","üzüm çekirdeği"], skin_types:["sensitive"], concerns:["moisture"], pregnancy:"safe", label:"Bitkisel sakinleştiriciler"},
  {match:["gliserin","glycerin"], skin_types:["dry","normal","oily"], concerns:["moisture"], pregnancy:"safe", label:"Gliserin"},
  {match:["çinko pca","zinc pca","çinko pirityon"], skin_types:["oily"], concerns:["acne"], pregnancy:"safe", label:"Çinko (Zinc PCA)"},
  {match:["vitamin e","tokoferol"], skin_types:["dry","normal"], concerns:["moisture","antiage"], pregnancy:"safe", label:"E Vitamini"},
  {match:["shea","karite","beeswax","hindistan cevizi","ayçiçek yağı"], skin_types:["dry"], concerns:["moisture"], pregnancy:"safe", label:"Bitkisel yağlar (Shea/Karite vb.)"},
  {match:["probiyotik","prebiyotik"], skin_types:["sensitive"], concerns:["moisture"], pregnancy:"safe", label:"Pro/Prebiyotik"},
  {match:["spf","mexoryl"], skin_types:[], concerns:["sun"], pregnancy:"consult", label:"Güneş filtresi"},
  {match:["hidrokinon","hydroquinone"], skin_types:["normal"], concerns:["pigment"], pregnancy:"avoid", label:"Hidrokinon"},
  {match:["kojik","arbutin","viniferine"], skin_types:["normal"], concerns:["pigment"], pregnancy:"consult", label:"Kojik Asit / Arbutin / Viniferine"},
  {match:["azelaik","azelaic"], skin_types:["oily","sensitive"], concerns:["acne","pigment"], pregnancy:"safe", label:"Azelaik Asit"},
  {match:["bakuchiol"], skin_types:["normal","sensitive"], concerns:["antiage"], pregnancy:"safe", label:"Bakuchiol"},
  {match:["üre","urea"], skin_types:["dry"], concerns:["moisture"], pregnancy:"safe", label:"Üre"},
  {match:["kaolin"], skin_types:["oily"], concerns:["acne"], pregnancy:"safe", label:"Kil (Kaolin)"},
  {match:["termal su","thermal water","mineralli su"], skin_types:["sensitive"], concerns:["moisture"], pregnancy:"safe", label:"Termal/Mineralli Su"},
  {match:["biotin"], skin_types:["normal"], concerns:["antiage"], pregnancy:"safe", label:"Biotin"},
  {match:["keratolytics"], skin_types:["oily"], concerns:["acne"], pregnancy:"consult", label:"Keratolitik ajanlar"},
];
const PREGNANCY_RANK={safe:1,consult:2,avoid:3};

function analyzeIngredients(actives){
  const skinSet=new Set(),concernSet=new Set(),matched=[];
  let worst=0,worstLabel="unknown";
  (actives||[]).forEach(raw=>{
    const low=" "+String(raw).toLowerCase()+" ";
    const hit=INGREDIENT_KB.find(k=>k.match.some(m=>low.includes(m)));
    if(hit){
      hit.skin_types.forEach(s=>skinSet.add(s));
      hit.concerns.forEach(c=>concernSet.add(c));
      matched.push({raw,label:hit.label,pregnancy:hit.pregnancy});
      if(PREGNANCY_RANK[hit.pregnancy]>worst){worst=PREGNANCY_RANK[hit.pregnancy];worstLabel=hit.pregnancy;}
    }
  });
  return{
    skin_types:[...skinSet], concerns:[...concernSet],
    pregnancy_safe: matched.length?worstLabel:"unknown",
    matched, unmatchedCount:(actives||[]).length-matched.length,
  };
}

// ══════════════════════════════════════════════════════════════
// İÇERİK ANALİZİ KUTUSU — havuz ürünlerinde kullanılan ortak bileşen
// ══════════════════════════════════════════════════════════════
function IngredientAnalysisBox({actives,onApply}){
  const [result,setResult]=useState(null);
  const PREG_META={safe:{l:"✓ Güvenli görünüyor",c:"#1E6B3E"},consult:{l:"⚠ Doktora danışılmalı",c:"#7D4700"},
    avoid:{l:"⛔ Kaçınılmalı",c:"#8B2E2E"},unknown:{l:"? Belirsiz",c:"#888"}};
  const analyze=()=>{
    const list=(Array.isArray(actives)?actives:String(actives||"").split(",")).map(s=>s.trim()).filter(Boolean);
    setResult(analyzeIngredients(list));
  };
  return(
    <div style={{marginTop:8}}>
      <button type="button" onClick={analyze}
        style={{background:"#EEF0FB",color:"#3949AB",border:"none",borderRadius:8,
          padding:"7px 12px",fontSize:11,fontWeight:600,cursor:"pointer"}}>
        🧪 İçerikten Otomatik Analiz Et
      </button>
      {result&&(
        <div style={{background:"#F8F9FA",borderRadius:10,padding:10,marginTop:8}}>
          {result.matched.length===0?(
            <div style={{fontSize:11,color:"#999"}}>Tanınan aktif içerik bulunamadı. Lütfen manuel seçim yapın.</div>
          ):(<>
            <div style={{fontSize:11,fontWeight:700,color:"#333",marginBottom:6}}>Analiz Sonucu (öneri):</div>
            <div style={{fontSize:11,color:"#555",marginBottom:3}}>
              <b>Cilt tipi:</b> {result.skin_types.length?result.skin_types.map(s=>SKIN_TYPE_LABELS[s]).join(", "):"—"}
            </div>
            <div style={{fontSize:11,color:"#555",marginBottom:3}}>
              <b>Endişe:</b> {result.concerns.length?result.concerns.map(c=>CONCERN_LABELS[c]).join(", "):"—"}
            </div>
            <div style={{fontSize:11,marginBottom:8,color:PREG_META[result.pregnancy_safe].c}}>
              <b style={{color:"#555"}}>Hamilelik:</b> {PREG_META[result.pregnancy_safe].l}
            </div>
            <div style={{fontSize:10,color:"#888",marginBottom:8,lineHeight:1.5}}>
              Tanınan: {result.matched.map(m=>m.label).join(", ")}
              {result.unmatchedCount>0&&` (+${result.unmatchedCount} tanınmayan içerik)`}
            </div>
            <button type="button" onClick={()=>onApply(result)}
              style={{background:"#1C1C1A",color:"#fff",border:"none",borderRadius:8,
                padding:"7px 12px",fontSize:11,fontWeight:600,cursor:"pointer"}}>
              Bu Önerileri Uygula
            </button>
          </>)}
          <div style={{fontSize:9,color:"#AAA",marginTop:8,lineHeight:1.5}}>
            İçerik listesine dayanan kural tabanlı bir öneridir, kesin bir tıbbi/farmasötik
            değerlendirme değildir. Özellikle hamilelik uygunluğunu eczacı/hekim onayıyla teyit edin.
          </div>
        </div>
      )}
    </div>
  );
}

const BRAND_COLORS = {
  "Avène":{"bg":"#D6E8F5","c":"#1A5276"},"Caudalie":{"bg":"#FDEEF8","c":"#6B2D5E"},
  "La Roche-Posay":{"bg":"#F5EEF8","c":"#4A235A"},"Bioderma":{"bg":"#FEF5E7","c":"#784212"},
  "Vichy":{"bg":"#E8F8F5","c":"#0E6655"},"Ducray":{"bg":"#EAFAF1","c":"#1E8449"},
  "CeraVe":{"bg":"#EBF5FB","c":"#1B2631"},"ISDIN":{"bg":"#FEFCE8","c":"#7D6608"},
  "La Rosée":{"bg":"#F5EDE8","c":"#5D4037"},"The Purest":{"bg":"#EEF0FB","c":"#1A237E"},"Dermoskin":{"bg":"#F1F8E9","c":"#1B5E20"},
};

// ══════════════════════════════════════════════════════════════
// ÜRÜN HAVUZU
// pregnancy_safe: "safe" | "avoid" | "consult" | "unknown"
// skin_types: ["oily","dry","sensitive","normal"]
// concerns: ["moisture","acne","pigment","antiage","sun"]
// ══════════════════════════════════════════════════════════════
const INITIAL_POOL = [
  // AVÈNE
  {id:1, barcode:"3282770203745",brand:"Avène",name:"Eau Thermale Temizleyici Jel",category:"temizleyici",usage:"both",basePrice:520,actives:["Avène Termal Suyu"],pairs_with:[2,5],certs:["dermo-test","hypoallergenic"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["sensitive","dry","normal"],concerns:["moisture"],desc:"Hassas ciltler için köpüksüz yumuşak temizleyici.",how_to:"Islak yüze 30 sn masaj, durulayın."},
  {id:2, barcode:"3282770073980",brand:"Avène",name:"Hydrance Intense Serum",category:"serum",usage:"both",basePrice:890,actives:["Hyalüronik Asit","Avène Termal Suyu"],pairs_with:[1,3],certs:["dermo-test","hypoallergenic"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["dry","sensitive","normal"],concerns:["moisture"],desc:"Çok katmanlı hyalüronik asit serumu.",how_to:"Temizleyiciden sonra 2–3 damla."},
  {id:3, barcode:"3282770204230",brand:"Avène",name:"Hydrance Rich Nemlendirici",category:"nemlendirici",usage:"both",basePrice:760,actives:["Avène Termal Suyu","Gliserin"],pairs_with:[1,2],certs:["dermo-test","hypoallergenic"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["dry","sensitive"],concerns:["moisture"],desc:"Kuru ve hassas ciltler için kalıcı nem.",how_to:"Serumdan sonra bezelye kadar."},
  {id:4, barcode:"3282770106459",brand:"Avène",name:"Cleanance Temizleyici Jel",category:"temizleyici",usage:"both",basePrice:485,actives:["Salisilik Asit","Avène Termal Suyu"],pairs_with:[13,20],certs:["dermo-test"],boycott:"temiz",pregnancy_safe:"consult",skin_types:["oily"],concerns:["acne"],desc:"Yağlı ve akneli ciltler için BHA içerikli jel.",how_to:"Yüze uygulayın, köpürtün, durulayın."},
  {id:5, barcode:"3282770023817",brand:"Avène",name:"Cicalfate+ Onarıcı Krem",category:"tedavi",usage:"both",basePrice:640,actives:["Çinko","Bakır","Sucralfate"],pairs_with:[1],certs:["dermo-test","hypoallergenic"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["sensitive","dry"],concerns:["moisture"],desc:"Tahriş ve hasar görmüş ciltler için bariyer onarıcı.",how_to:"Temiz cilde ince tabaka."},
  // CAUDALIE
  {id:6, barcode:"3522930001721",brand:"Caudalie",name:"Vinoclean Micellar Su",category:"temizleyici",usage:"both",basePrice:545,actives:["Üzüm Suyu","Polifenol"],pairs_with:[7,8],certs:["vegan","dermo-test"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["normal","dry","sensitive"],concerns:["moisture"],desc:"Makyaj ve kirleticileri nazikçe temizler.",how_to:"Pamukla silin, durulamayın."},
  {id:7, barcode:"3522930002285",brand:"Caudalie",name:"Vinoperfect Işıltı Serumu",category:"serum",usage:"am",basePrice:1250,actives:["Viniferine","Traneksamik Asit","Niasinamid"],pairs_with:[6,8],certs:["vegan","dermo-test"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["normal","dry","oily"],concerns:["pigment"],desc:"Leke karşıtı, ton eşitleme sabah serumu.",how_to:"Sabah 3 damla."},
  {id:8, barcode:"3522930001813",brand:"Caudalie",name:"Vinosource SOS Krem",category:"nemlendirici",usage:"both",basePrice:890,actives:["Üzüm Suyu","Gliserin","Resveratrol"],pairs_with:[6,7],certs:["vegan","dermo-test"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["dry","sensitive","normal"],concerns:["moisture","antiage"],desc:"Susuz ciltler için yoğun nem.",how_to:"Serumdan sonra sabah-akşam."},
  {id:9, barcode:"3522930004395",brand:"Caudalie",name:"Beauty Elixir Sprey",category:"tonik",usage:"both",basePrice:780,actives:["Gül Suyu","Polifenol"],pairs_with:[7,8],certs:["vegan"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["normal","dry","oily"],concerns:["moisture"],desc:"Günün her saati canlandırıcı sprey.",how_to:"30 cm uzaktan 2–3 kez sık."},
  // DUCRAY
  {id:10,barcode:"3401396818759",brand:"Ducray",name:"Keracnyl Temizleyici Jel",category:"temizleyici",usage:"both",basePrice:430,actives:["Niasinamid","Çinko"],pairs_with:[11,12],certs:["dermo-test"],boycott:"bilinmiyor",pregnancy_safe:"safe",skin_types:["oily"],concerns:["acne"],desc:"Akneye yatkın yağlı ciltler için.",how_to:"Islak yüze uygulayın, durulayın."},
  {id:11,barcode:"3401399103697",brand:"Ducray",name:"Melascreen UV SPF50+",category:"spf",usage:"am",basePrice:920,actives:["SPF50+"],pairs_with:[10,12],certs:["dermo-test","spf"],boycott:"bilinmiyor",pregnancy_safe:"consult",skin_types:["oily","normal"],concerns:["pigment","sun"],desc:"Güneş lekelerine karşı yüksek koruma.",how_to:"Sabah son adım, 20 dk önce."},
  {id:12,barcode:"3401399234040",brand:"Ducray",name:"Ictyane Nemlendirici Krem",category:"nemlendirici",usage:"both",basePrice:510,actives:["Shea Yağı","Gliserin"],pairs_with:[10],certs:["dermo-test","hypoallergenic"],boycott:"bilinmiyor",pregnancy_safe:"safe",skin_types:["dry","sensitive"],concerns:["moisture"],desc:"Kuru ve atopiye yatkın ciltler.",how_to:"Banyo sonrası ıslak cilde."},
  // LA ROCHE-POSAY
  {id:13,barcode:"3337872413361",brand:"La Roche-Posay",name:"Effaclar Jel Temizleyici",category:"temizleyici",usage:"both",basePrice:465,actives:["Niasinamid","Salisilik Asit"],pairs_with:[14,15],certs:["dermo-test","fragrance-free"],boycott:"temiz",pregnancy_safe:"consult",skin_types:["oily"],concerns:["acne"],desc:"Yağlı ve akneli ciltler için arındırıcı jel.",how_to:"Yüze uygulayın, durulayın."},
  {id:14,barcode:"3337872413378",brand:"La Roche-Posay",name:"Hyalu B5 Serum",category:"serum",usage:"both",basePrice:980,actives:["Hyalüronik Asit","Vitamin B5"],pairs_with:[13,15],certs:["dermo-test","hypoallergenic"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["dry","normal","sensitive"],concerns:["moisture"],desc:"Bariyer onarıcı nem serumu.",how_to:"4–5 damla, masaj yapın."},
  {id:15,barcode:"3337875520676",brand:"La Roche-Posay",name:"Toleriane Double Repair",category:"nemlendirici",usage:"both",basePrice:720,actives:["Seramid","Niasinamid","Prebiyotik"],pairs_with:[13,14],certs:["dermo-test","hypoallergenic","fragrance-free"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["sensitive","dry","normal"],concerns:["moisture"],desc:"Hassas ciltler için çift onarımlı.",how_to:"Sabah ve akşam."},
  {id:16,barcode:"3337875486124",brand:"La Roche-Posay",name:"Anthelios UV Mune SPF50+",category:"spf",usage:"am",basePrice:890,actives:["SPF50+","Mexoryl 400"],pairs_with:[13,7],certs:["dermo-test","spf"],boycott:"temiz",pregnancy_safe:"consult",skin_types:["sensitive","normal","dry"],concerns:["sun"],desc:"Hassas ciltler için gelişmiş UV.",how_to:"Sabah son adım."},
  // BIODERMA
  {id:17,barcode:"3401346691320",brand:"Bioderma",name:"Sensibio H2O Micellar Su",category:"temizleyici",usage:"both",basePrice:380,actives:["Salatalık Özü"],pairs_with:[18,19],certs:["dermo-test","hypoallergenic","fragrance-free"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["sensitive","normal","dry"],concerns:["moisture"],desc:"Hassas ciltler için micellar su.",how_to:"Pamukla silin, durulamayın."},
  {id:18,barcode:"3401346398690",brand:"Bioderma",name:"Hydrabio Serum",category:"serum",usage:"both",basePrice:760,actives:["Hyalüronik Asit","Vitamin PP"],pairs_with:[17,19],certs:["dermo-test","hypoallergenic"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["dry","normal"],concerns:["moisture"],desc:"Dehidrate ciltler için serum.",how_to:"Sabah-akşam uygulayın."},
  {id:19,barcode:"3401346691115",brand:"Bioderma",name:"Sebium Pore Refiner",category:"tedavi",usage:"am",basePrice:620,actives:["Niasinamid","Çinko"],pairs_with:[17],certs:["dermo-test"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["oily"],concerns:["acne"],desc:"Gözenek azaltıcı, mat bırakıcı.",how_to:"Sabah nemlendirici öncesi."},
  // VİCHY
  {id:20,barcode:"3337875587075",brand:"Vichy",name:"Normaderm Phytosolution Jel",category:"temizleyici",usage:"both",basePrice:445,actives:["Salisilik Asit","Probiyotik"],pairs_with:[21,22],certs:["dermo-test"],boycott:"temiz",pregnancy_safe:"consult",skin_types:["oily"],concerns:["acne"],desc:"Karma ve yağlı ciltler için probiyotik jel.",how_to:"Islak yüze uygulayın, durulayın."},
  {id:21,barcode:"3337875694285",brand:"Vichy",name:"Minéral 89 Booster",category:"serum",usage:"both",basePrice:680,actives:["Hyalüronik Asit","Vichy Mineralli Su"],pairs_with:[20,22],certs:["dermo-test","fragrance-free"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["normal","dry","sensitive"],concerns:["moisture"],desc:"%89 mineralll su içerikli güçlendirici.",how_to:"Temizleyiciden sonra 3–4 damla."},
  {id:22,barcode:"3337875501283",brand:"Vichy",name:"Capital Soleil UV-Age SPF50+",category:"spf",usage:"am",basePrice:810,actives:["SPF50+","Niasinamid"],pairs_with:[20,21],certs:["dermo-test","spf"],boycott:"temiz",pregnancy_safe:"consult",skin_types:["normal","oily"],concerns:["sun","antiage"],desc:"Anti-aging özellikli günlük SPF.",how_to:"Sabah son adım."},
  // CERAVE
  {id:23,barcode:"3606000594340",brand:"CeraVe",name:"Hydrating Cleanser",category:"temizleyici",usage:"both",basePrice:410,actives:["Seramid","Hyalüronik Asit"],pairs_with:[24,25],certs:["fragrance-free","hypoallergenic"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["dry","sensitive","normal"],concerns:["moisture"],desc:"Seramid içerikli nazik temizleyici.",how_to:"Islak yüze uygulayın, durulayın."},
  {id:24,barcode:"3606000594357",brand:"CeraVe",name:"Moisturizing Cream",category:"nemlendirici",usage:"both",basePrice:520,actives:["Seramid","Hyalüronik Asit","Niasinamid"],pairs_with:[23,25],certs:["fragrance-free","hypoallergenic"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["dry","sensitive","normal"],concerns:["moisture"],desc:"24 saat nem kilitleyen krem.",how_to:"Sabah ve akşam."},
  {id:25,barcode:"3606000595309",brand:"CeraVe",name:"Resurfacing Retinol Serum",category:"serum",usage:"pm",basePrice:890,actives:["Retinol","Niasinamid","Seramid"],pairs_with:[23,24],certs:["fragrance-free"],boycott:"temiz",pregnancy_safe:"avoid",skin_types:["normal","oily"],concerns:["antiage","acne"],desc:"Retinol içerikli yenileme serumu.",how_to:"Sadece akşam, kuru cilde."},
  // ISDIN
  {id:26,barcode:"8429420188649",brand:"ISDIN",name:"Fotoprotector Fusion Water SPF50",category:"spf",usage:"am",basePrice:760,actives:["SPF50"],pairs_with:[7,14],certs:["dermo-test","spf"],boycott:"bilinmiyor",pregnancy_safe:"consult",skin_types:["oily","normal"],concerns:["sun"],desc:"Su gibi hafif dokulu SPF.",how_to:"Sabah son adım."},
  {id:27,barcode:"8429420249907",brand:"ISDIN",name:"Eryfotona Actinica SPF100+",category:"spf",usage:"am",basePrice:1100,actives:["SPF100+","DNA Onarım Enzimleri"],pairs_with:[13,20],certs:["dermo-test","spf"],boycott:"bilinmiyor",pregnancy_safe:"consult",skin_types:["sensitive","normal"],concerns:["sun"],desc:"Maksimum UV koruması.",how_to:"Nokta nokta sürün."},
  // LA ROSÉE
  {id:28,barcode:"3760168131018",brand:"La Rosée",name:"Eau Micellaire Douce",category:"temizleyici",usage:"both",basePrice:340,actives:["Hindistan Cevizi Yağı","Gliserin"],pairs_with:[29,30],certs:["natural","vegan","fragrance-free"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["sensitive","dry","normal"],concerns:["moisture"],desc:"Doğal içerikli, nazik micellar su.",how_to:"Pamukla silin, durulamayın."},
  {id:29,barcode:"3760168131025",brand:"La Rosée",name:"Crème Hydratante Universelle",category:"nemlendirici",usage:"both",basePrice:380,actives:["Hindistan Cevizi Yağı","Shea Yağı","Gliserin"],pairs_with:[28,30],certs:["natural","vegan"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["dry","sensitive","normal"],concerns:["moisture"],desc:"Tüm cilt tipleri için çok amaçlı nemlendirici.",how_to:"Yüz, el, vücuda uygulayın."},
  {id:30,barcode:"3760168131032",brand:"La Rosée",name:"Sérum Éclat Vitamine C",category:"serum",usage:"am",basePrice:490,actives:["Vitamin C","Gül Suyu"],pairs_with:[28,29],certs:["natural","vegan"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["normal","dry"],concerns:["pigment","antiage"],desc:"Doğal Vitamin C ile aydınlatma serumu.",how_to:"Sabah 2–3 damla."},
  {id:31,barcode:"3760168131049",brand:"La Rosée",name:"Beurre Nourrissant Corps",category:"nemlendirici",usage:"both",basePrice:290,actives:["Shea Yağı","Karboksimetil Beta Glukan"],pairs_with:[28],certs:["natural","vegan"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["dry","sensitive"],concerns:["moisture"],desc:"Vücut için yoğun besleyici tereyağı.",how_to:"Banyo sonrası cilde yedirin."},
  // THE PUREST
  {id:32,barcode:"8697939390012",brand:"The Purest",name:"Salicylic Acid Tonik",category:"tonik",usage:"pm",basePrice:395,actives:["Salisilik Asit %1","Niasinamid"],pairs_with:[33,34],certs:["dermo-test"],boycott:"temiz",pregnancy_safe:"consult",skin_types:["oily"],concerns:["acne"],desc:"Gözenek açıcı, akne karşıtı tonik.",how_to:"Pamukla akşam tonikleme yapın."},
  {id:33,barcode:"8697939390029",brand:"The Purest",name:"Hyaluronic Acid Serum",category:"serum",usage:"both",basePrice:345,actives:["Hyalüronik Asit","B5 Vitamini"],pairs_with:[32,34,23],certs:["dermo-test","fragrance-free"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["dry","normal","sensitive"],concerns:["moisture"],desc:"Çok ağırlıklı hyalüronik asit serumu.",how_to:"Temizleyiciden sonra 3–4 damla."},
  {id:34,barcode:"8697939390036",brand:"The Purest",name:"Niacinamide %10 Serum",category:"serum",usage:"both",basePrice:295,actives:["Niasinamid %10","Çinko"],pairs_with:[32,33],certs:["dermo-test","fragrance-free"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["oily","normal"],concerns:["acne","pigment"],desc:"Gözenek sıkılaştırıcı, ton eşitleyici.",how_to:"Temizleyiciden sonra 2–3 damla."},
  {id:35,barcode:"8697939390043",brand:"The Purest",name:"Soothing Centella Krem",category:"nemlendirici",usage:"both",basePrice:350,actives:["Centella Asiatica","Madecassoside"],pairs_with:[33],certs:["dermo-test","fragrance-free","hypoallergenic"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["sensitive","dry"],concerns:["moisture"],desc:"Centella ile yatıştırıcı, onarıcı krem.",how_to:"Sabah ve akşam uygulayın."},

  // ── AVÈNE (genişletilmiş) ────────────────────────────────────
  {id:36,barcode:"3282770207491",brand:"Avène",name:"Tolérance Extrême Krem",category:"nemlendirici",usage:"both",basePrice:680,actives:["Avène Termal Suyu","Squalane"],pairs_with:[1,5],certs:["dermo-test","hypoallergenic","fragrance-free"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["sensitive"],concerns:["moisture"],desc:"Aşırı hassas ve reaktif ciltler için ultra yatıştırıcı krem.",how_to:"İnce tabaka halinde uygulayın."},
  {id:37,barcode:"3282770207507",brand:"Avène",name:"Antirougeurs Fort Krem",category:"nemlendirici",usage:"both",basePrice:720,actives:["Avène Termal Suyu","Ruscus Özü"],pairs_with:[1],certs:["dermo-test","hypoallergenic"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["sensitive"],concerns:["moisture"],desc:"Kızarıklık ve rozase eğilimli ciltler için bakım.",how_to:"Sabah ve akşam temiz cilde uygulayın."},
  {id:38,barcode:"3282770104387",brand:"Avène",name:"TriAcnéal Expert Krem",category:"tedavi",usage:"pm",basePrice:580,actives:["Retinaldehid","Diolenyl","Glikolik Asit"],pairs_with:[1,3],certs:["dermo-test"],boycott:"temiz",pregnancy_safe:"avoid",skin_types:["oily","normal"],concerns:["acne","antiage"],desc:"Akne sonrası izler ve yağlı cilt için aktif bakım.",how_to:"Akşam nokta nokta uygulayın."},
  {id:39,barcode:"3282770208733",brand:"Avène",name:"PhysioLift Günlük Bakım",category:"nemlendirici",usage:"am",basePrice:850,actives:["Retinal","Kaolin","Hyalüronik Asit"],pairs_with:[2,3],certs:["dermo-test"],boycott:"temiz",pregnancy_safe:"consult",skin_types:["dry","normal"],concerns:["antiage"],desc:"Yaşlanma karşıtı dolgunlaştırıcı günlük bakım.",how_to:"Sabah masaj yaparak uygulayın."},
  {id:40,barcode:"3282770073256",brand:"Avène",name:"XeraCalm A.D Lipid Krem",category:"nemlendirici",usage:"both",basePrice:620,actives:["I-modulia","Avène Termal Suyu"],pairs_with:[1,5],certs:["dermo-test","hypoallergenic","fragrance-free"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["sensitive","dry"],concerns:["moisture"],desc:"Atopi ve kaşıntı eğilimli kuru ciltler için.",how_to:"Banyo sonrası nemliyken uygulayın."},
  {id:41,barcode:"3282770208429",brand:"Avène",name:"Cleanance COMEDOMED Krem",category:"tedavi",usage:"both",basePrice:490,actives:["Niasinamid","Avène Termal Suyu","Retinaldehid"],pairs_with:[4],certs:["dermo-test"],boycott:"temiz",pregnancy_safe:"consult",skin_types:["oily"],concerns:["acne"],desc:"Siyah nokta ve mikrokist karşıtı krem.",how_to:"Sabah ve akşam temizleyici sonrası."},
  {id:42,barcode:"3282770022513",brand:"Avène",name:"Solaire SPF50+ Fluid",category:"spf",usage:"am",basePrice:520,actives:["SPF50+","Avène Termal Suyu"],pairs_with:[1,3],certs:["dermo-test","hypoallergenic","spf"],boycott:"temiz",pregnancy_safe:"consult",skin_types:["sensitive","normal"],concerns:["sun"],desc:"Hassas ciltler için hafif dokulu güneş fluid.",how_to:"Sabah son adım, bol miktarda uygulayın."},

  // ── CAUDALIE (genişletilmiş) ─────────────────────────────────
  {id:43,barcode:"3522930024317",brand:"Caudalie",name:"Vinosource-Hydra Serum",category:"serum",usage:"both",basePrice:980,actives:["Hyalüronik Asit","Grape Water","Polifenol"],pairs_with:[6,8],certs:["vegan","dermo-test"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["dry","normal"],concerns:["moisture"],desc:"Yoğun nem için Vinosource hattının serumu.",how_to:"Nemlendirici öncesi 3–4 damla."},
  {id:44,barcode:"3522930019481",brand:"Caudalie",name:"Vinopure Matlaştırıcı Tonik",category:"tonik",usage:"both",basePrice:580,actives:["Salisilik Asit","Gül Suyu","Niasinamid"],pairs_with:[6,7],certs:["vegan"],boycott:"temiz",pregnancy_safe:"consult",skin_types:["oily"],concerns:["acne","pigment"],desc:"Yağlı cilt için gözenek sıkılaştırıcı tonik.",how_to:"Pamukla silin ya da deri üzerine sıkın."},
  {id:45,barcode:"3522930018279",brand:"Caudalie",name:"Resveratrol Lift Night Cream",category:"nemlendirici",usage:"pm",basePrice:1350,actives:["Resveratrol","Vitis Vinifera","Hyalüronik Asit"],pairs_with:[6,9],certs:["vegan","dermo-test"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["normal","dry"],concerns:["antiage"],desc:"Gece boyunca sıkılaştırıcı ve yenileyici bakım.",how_to:"Akşam bol miktarda uygulayın."},
  {id:46,barcode:"3522930019405",brand:"Caudalie",name:"Vinopure Blemish Control Serum",category:"serum",usage:"both",basePrice:890,actives:["Salisilik Asit","Niasinamid","Resveratrol"],pairs_with:[44,6],certs:["vegan","dermo-test"],boycott:"temiz",pregnancy_safe:"consult",skin_types:["oily"],concerns:["acne"],desc:"Akne ve leke eğilimli ciltler için serum.",how_to:"Temizleyiciden sonra 2–3 damla."},
  {id:47,barcode:"3522930018989",brand:"Caudalie",name:"Polyphenol C15 Anti-Wrinkle Serum",category:"serum",usage:"am",basePrice:1150,actives:["Vitamin C","Resveratrol","Polifenol"],pairs_with:[8,9],certs:["vegan","dermo-test"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["normal","dry"],concerns:["antiage","pigment"],desc:"Vitamin C ve resveratrol ile yaşlanma karşıtı serum.",how_to:"Sabah 4–5 damla."},
  {id:48,barcode:"3522930003121",brand:"Caudalie",name:"Vinoclean Cleansing Oil",category:"temizleyici",usage:"pm",basePrice:620,actives:["Üzüm Çekirdeği Yağı","Ayçiçek Yağı"],pairs_with:[6,7],certs:["vegan"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["dry","normal","sensitive"],concerns:["moisture"],desc:"Çift temizleme için yağ bazlı makyaj çözücü.",how_to:"Kuru cilde uygulayın, su ekleyip masaj yapın."},

  // ── DUCRAY (genişletilmiş) ───────────────────────────────────
  {id:49,barcode:"3282770109320",brand:"Ducray",name:"Sensinol Arındırıcı Şampuan",category:"tedavi",usage:"both",basePrice:285,actives:["Piroctone Olamine","Sfingozin"],pairs_with:[],certs:["dermo-test"],boycott:"bilinmiyor",pregnancy_safe:"consult",skin_types:["sensitive"],concerns:["moisture"],desc:"Hassas kafa derisi için yatıştırıcı şampuan.",how_to:"Islak saça uygulayın, 3 dk bekleyin, durulayın."},
  {id:50,barcode:"3282770013436",brand:"Ducray",name:"Creastim Dökülme Karşıtı Losyon",category:"tedavi",usage:"both",basePrice:450,actives:["Adenos-H","Mangan","Çinko"],pairs_with:[],certs:["dermo-test"],boycott:"bilinmiyor",pregnancy_safe:"consult",skin_types:["normal","oily"],concerns:["moisture"],desc:"Erkek ve kadın tipi dökülmeye karşı losyon.",how_to:"Haftada 2–3 kez kafa derisine uygulayın."},
  {id:51,barcode:"3282770076578",brand:"Ducray",name:"Kelual DS Kepek Şampuanı",category:"tedavi",usage:"both",basePrice:310,actives:["Cyclopyroxyolamine","Keratolytics","Klimazol"],pairs_with:[],certs:["dermo-test"],boycott:"bilinmiyor",pregnancy_safe:"consult",skin_types:["oily","normal"],concerns:["acne"],desc:"Dirençli kepek ve seboreik dermatit şampuanı.",how_to:"Haftada 2 kez kullanın, köpürtün ve bekleyin."},
  {id:52,barcode:"3282770076295",brand:"Ducray",name:"Nutricerat Yoğun Bakım Maskesi",category:"maske",usage:"both",basePrice:380,actives:["Ceramid","Shea Yağı","Karite"],pairs_with:[12],certs:["dermo-test"],boycott:"bilinmiyor",pregnancy_safe:"safe",skin_types:["dry","sensitive"],concerns:["moisture"],desc:"Aşırı kuru ve atopiye yatkın saçlar için bakım maskesi.",how_to:"Şampuan sonrası 5 dk uygulayın, durulayın."},
  {id:53,barcode:"3282770011050",brand:"Ducray",name:"Diaseptyl Yara İyileştirici Sprey",category:"tedavi",usage:"both",basePrice:220,actives:["Biseptol","Benzalkonium"],pairs_with:[],certs:["dermo-test"],boycott:"bilinmiyor",pregnancy_safe:"consult",skin_types:["sensitive","normal"],concerns:["moisture"],desc:"Küçük yara ve ciltte tahriş için antiseptik sprey.",how_to:"Günde 2–3 kez hafif uzaklıktan sıkın."},
  {id:54,barcode:"3282770076561",brand:"Ducray",name:"Squanorm Yağlı Kepek Şampuanı",category:"tedavi",usage:"both",basePrice:295,actives:["Cyclopyroxyolamine","Salisilik Asit"],pairs_with:[],certs:["dermo-test"],boycott:"bilinmiyor",pregnancy_safe:"consult",skin_types:["oily"],concerns:["acne"],desc:"Yağlı kepek için güçlü etkili şampuan.",how_to:"Haftada 2 kez kullanın."},
  {id:55,barcode:"3282770076547",brand:"Ducray",name:"Kertyol PSO Şampuan",category:"tedavi",usage:"both",basePrice:340,actives:["Keratolytics","Salisilik Asit","Çinko Pirityon"],pairs_with:[],certs:["dermo-test"],boycott:"bilinmiyor",pregnancy_safe:"consult",skin_types:["oily","normal"],concerns:["acne"],desc:"Psöriyazis ve kepek için keratoletik şampuan.",how_to:"Haftada 2–3 kez, bekleterek kullanın."},

  // ── LA ROCHE-POSAY (genişletilmiş) ──────────────────────────
  {id:56,barcode:"3337875597500",brand:"La Roche-Posay",name:"Effaclar Duo+ Krem",category:"tedavi",usage:"both",basePrice:580,actives:["Niasinamid","LHA","Proksilin"],pairs_with:[13,15],certs:["dermo-test"],boycott:"temiz",pregnancy_safe:"consult",skin_types:["oily"],concerns:["acne"],desc:"Akne lekeleri ve gözenek sıkılaştırıcı krem.",how_to:"Sabah ve akşam sorunlu bölgelere uygulayın."},
  {id:57,barcode:"3337875597517",brand:"La Roche-Posay",name:"Lipikar Baume AP+ Krem",category:"nemlendirici",usage:"both",basePrice:490,actives:["Shea Yağı","Niasinamid","Glycerin"],pairs_with:[15],certs:["dermo-test","hypoallergenic","fragrance-free"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["dry","sensitive"],concerns:["moisture"],desc:"Atopik ve kuru vücut derisi için yoğun bakım.",how_to:"Banyo sonrası tüm vücuda uygulayın."},
  {id:58,barcode:"3337875597524",brand:"La Roche-Posay",name:"Pure Vitamin C10 Serum",category:"serum",usage:"am",basePrice:920,actives:["Vitamin C %10","Salisilik Asit","Nörotensid"],pairs_with:[13,15],certs:["dermo-test"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["normal","oily"],concerns:["pigment","antiage"],desc:"Saf Vitamin C ile parlaklık ve leke karşıtı serum.",how_to:"Sabah 4–5 damla, SPF öncesi."},
  {id:59,barcode:"3337875597531",brand:"La Roche-Posay",name:"Mela B3 Leke Karşıtı Serum",category:"serum",usage:"both",basePrice:1050,actives:["Niasinamid","Traneksamik Asit","LHA"],pairs_with:[13,16],certs:["dermo-test"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["normal","oily","dry"],concerns:["pigment"],desc:"Isıratlı leke ve hiperpigmentasyon serumu.",how_to:"Sabah ve akşam 4–5 damla."},
  {id:60,barcode:"3337875597548",brand:"La Roche-Posay",name:"Nutritic Intense Krem",category:"nemlendirici",usage:"both",basePrice:680,actives:["Nörotensid","Shea Yağı","Gliserin"],pairs_with:[14,15],certs:["dermo-test","hypoallergenic"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["dry","sensitive"],concerns:["moisture"],desc:"Aşırı kuru ve hassas ciltler için besleyici krem.",how_to:"Sabah ve akşam bol miktarda."},
  {id:61,barcode:"3337875597555",brand:"La Roche-Posay",name:"Effaclar Micellar Su",category:"temizleyici",usage:"both",basePrice:320,actives:["Niasinamid","Çinko PCA"],pairs_with:[13,56],certs:["dermo-test","fragrance-free"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["oily"],concerns:["acne"],desc:"Yağlı ve akneli ciltler için yağ ve kir giderici micellar.",how_to:"Pamukla silin, durulamayın."},
  {id:62,barcode:"3337875486612",brand:"La Roche-Posay",name:"Cicaplast Baume B5+ Sprey",category:"tedavi",usage:"both",basePrice:420,actives:["Pantenol","Bisabolol","Manganez"],pairs_with:[15,57],certs:["dermo-test","hypoallergenic","fragrance-free"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["sensitive","dry"],concerns:["moisture"],desc:"Tahriş olan bölgelere pratik sprey uygulaması.",how_to:"10–15 cm uzaktan sıkın, kanatın."},

  // ── LA ROSÉE (genişletilmiş) ─────────────────────────────────
  {id:63,barcode:"3760168131056",brand:"La Rosée",name:"Gel Nettoyant Doux Visage",category:"temizleyici",usage:"both",basePrice:280,actives:["Hindistan Cevizi Yağı","Aloe Vera"],pairs_with:[29,30],certs:["natural","vegan","fragrance-free"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["normal","sensitive","dry"],concerns:["moisture"],desc:"Doğal içerikli yüz yıkama jeli.",how_to:"Islak yüze uygulayın, durulayın."},
  {id:64,barcode:"3760168131063",brand:"La Rosée",name:"Huile Démaquillante Douce",category:"temizleyici",usage:"pm",basePrice:310,actives:["Ayçiçek Yağı","Vitamin E"],pairs_with:[28,63],certs:["natural","vegan"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["dry","sensitive","normal"],concerns:["moisture"],desc:"Nazik çifte temizleme yağı.",how_to:"Kuru cilde uygulayın, su ekleyip emülsifiye edin."},
  {id:65,barcode:"3760168131070",brand:"La Rosée",name:"Stick Lèvres Hydratant",category:"tedavi",usage:"both",basePrice:120,actives:["Shea Yağı","Beeswax","Vitamin E"],pairs_with:[],certs:["natural"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["sensitive","dry","normal"],concerns:["moisture"],desc:"Yoğun nemlendirici dudak bakım çubuğu.",how_to:"Gerektiğinde dudaklara uygulayın."},
  {id:66,barcode:"3760168131087",brand:"La Rosée",name:"SPF50 Crème Solaire Visage",category:"spf",usage:"am",basePrice:380,actives:["SPF50","Titanium Dioxide","Zinc Oxide"],pairs_with:[29,63],certs:["natural","vegan","spf"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["sensitive","normal","dry"],concerns:["sun"],desc:"Mineral filtreli doğal güneş koruyucu.",how_to:"Güneşe çıkmadan 20 dk önce uygulayın."},
  {id:67,barcode:"3760168131094",brand:"La Rosée",name:"Crème Apaisante Corps",category:"nemlendirici",usage:"both",basePrice:260,actives:["Aloe Vera","Gliserin","Calendula"],pairs_with:[28,31],certs:["natural","vegan","fragrance-free"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["sensitive","dry"],concerns:["moisture"],desc:"Vücut için yatıştırıcı ve nemlendirici krem.",how_to:"Temiz cilde bolca uygulayın."},

  // ── CERAVE (genişletilmiş) ───────────────────────────────────
  {id:68,barcode:"3606000594371",brand:"CeraVe",name:"SA Smoothing Cleanser",category:"temizleyici",usage:"both",basePrice:440,actives:["Salisilik Asit","Seramid","Laktik Asit"],pairs_with:[24,25],certs:["fragrance-free"],boycott:"temiz",pregnancy_safe:"consult",skin_types:["oily","normal"],concerns:["acne"],desc:"Pürüzsüzleştirici salisilik asit temizleyici.",how_to:"Islak cilde uygulayın, durulayın."},
  {id:69,barcode:"3606000594388",brand:"CeraVe",name:"Foaming Facial Cleanser",category:"temizleyici",usage:"both",basePrice:420,actives:["Seramid","Niasinamid","Hyalüronik Asit"],pairs_with:[24],certs:["fragrance-free","hypoallergenic"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["oily","normal"],concerns:["moisture","acne"],desc:"Normal ve yağlı ciltler için köpüklü temizleyici.",how_to:"Islak yüze uygulayın, durulayın."},
  {id:70,barcode:"3606000594395",brand:"CeraVe",name:"Eye Repair Cream",category:"goz",usage:"both",basePrice:580,actives:["Hyalüronik Asit","Seramid","Niasinamid"],pairs_with:[24],certs:["fragrance-free","hypoallergenic"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["sensitive","dry","normal"],concerns:["moisture","antiage"],desc:"Göz çevresi için onarıcı ve nemlendirici krem.",how_to:"Sabah ve akşam parmak uçlarıyla yedirin."},
  {id:71,barcode:"3606000594401",brand:"CeraVe",name:"Blemish Control Gel",category:"tedavi",usage:"both",basePrice:390,actives:["Salisilik Asit","Niasinamid","Çinko PCA"],pairs_with:[68,69],certs:["fragrance-free"],boycott:"temiz",pregnancy_safe:"consult",skin_types:["oily"],concerns:["acne"],desc:"Akne ve kabarıklık kontrolü için jel.",how_to:"Temizleyici sonrası sorunlu bölgelere."},
  {id:72,barcode:"3606000594418",brand:"CeraVe",name:"AM Facial Moisturising Lotion SPF25",category:"spf",usage:"am",basePrice:480,actives:["SPF25","Seramid","Niasinamid","Hyalüronik Asit"],pairs_with:[23,24],certs:["fragrance-free","spf"],boycott:"temiz",pregnancy_safe:"consult",skin_types:["normal","dry"],concerns:["moisture","sun"],desc:"Seramid içerikli nem ve güneş koruması bir arada.",how_to:"Sabah son adım olarak uygulayın."},

  // ── THE PUREST (genişletilmiş) ───────────────────────────────
  {id:73,barcode:"8697939390050",brand:"The Purest",name:"Retinol 0.1% Serum",category:"serum",usage:"pm",basePrice:420,actives:["Retinol %0.1","Hyalüronik Asit"],pairs_with:[33,35],certs:["dermo-test","fragrance-free"],boycott:"temiz",pregnancy_safe:"avoid",skin_types:["normal","oily"],concerns:["antiage","acne"],desc:"Başlangıç seviyesi retinol serumu, cilt yenileme.",how_to:"Akşam kuru cilde 2–3 damla."},
  {id:74,barcode:"8697939390067",brand:"The Purest",name:"AHA+BHA Exfoliant Tonik",category:"tonik",usage:"pm",basePrice:365,actives:["Glikolik Asit","Salisilik Asit","Laktik Asit"],pairs_with:[32,34],certs:["dermo-test","fragrance-free"],boycott:"temiz",pregnancy_safe:"consult",skin_types:["oily","normal"],concerns:["acne","pigment"],desc:"Kimyasal peeling etkili çoklu asit tonik.",how_to:"Haftada 2–3 kez pamukla uygulayın."},
  {id:75,barcode:"8697939390074",brand:"The Purest",name:"Vitamin C+E Brightening Serum",category:"serum",usage:"am",basePrice:390,actives:["Vitamin C","Vitamin E","Ferulik Asit"],pairs_with:[33,35],certs:["dermo-test","fragrance-free"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["normal","dry"],concerns:["pigment","antiage"],desc:"Vitamin C+E ile aydınlatma ve antioksidan koruması.",how_to:"Sabah 3–4 damla."},
  {id:76,barcode:"8697939390081",brand:"The Purest",name:"Collagen Booster Krem",category:"nemlendirici",usage:"both",basePrice:410,actives:["Kolajen","Peptid","Hyalüronik Asit"],pairs_with:[33,75],certs:["dermo-test"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["normal","dry"],concerns:["antiage","moisture"],desc:"Kolajen ve peptid ile sıkılaştırıcı nemlendirici.",how_to:"Sabah ve akşam uygulayın."},

  // ── DERMOSKIN ────────────────────────────────────────────────
  {id:77,barcode:"8690104050260",brand:"Dermoskin",name:"Benzoil Peroksit %5 Jel",category:"tedavi",usage:"pm",basePrice:185,actives:["Benzoil Peroksit %5"],pairs_with:[],certs:["dermo-test"],boycott:"temiz",pregnancy_safe:"avoid",skin_types:["oily"],concerns:["acne"],desc:"Akne tedavisinde kullanılan antiseptik jel.",how_to:"Geceleri sorunlu bölgeye nokta uygulayın."},
  {id:78,barcode:"8690104050277",brand:"Dermoskin",name:"Anti Hair Loss Şampuan",category:"tedavi",usage:"both",basePrice:225,actives:["Biotin","Keratin","Çinko"],pairs_with:[],certs:["dermo-test"],boycott:"temiz",pregnancy_safe:"consult",skin_types:["oily","normal"],concerns:["moisture"],desc:"Saç dökülmesine karşı güçlendirici şampuan.",how_to:"Islak saça uygulayın, 3 dk bekleyin, durulayın."},
  {id:79,barcode:"8690104050284",brand:"Dermoskin",name:"Multivitamin Nemlendirici Krem",category:"nemlendirici",usage:"both",basePrice:210,actives:["Vitamin A","Vitamin E","Vitamin C","Çinko"],pairs_with:[],certs:["dermo-test"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["normal","dry"],concerns:["moisture","antiage"],desc:"Çok vitaminli günlük bakım kremi.",how_to:"Sabah ve akşam temiz cilde uygulayın."},
  {id:80,barcode:"8690104050291",brand:"Dermoskin",name:"Güneş Koruyucu SPF50+ Krem",category:"spf",usage:"am",basePrice:195,actives:["SPF50+","Titanium Dioxide"],pairs_with:[79],certs:["dermo-test","spf"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["normal","sensitive"],concerns:["sun"],desc:"Günlük kullanım için SPF50+ güneş koruyucu.",how_to:"Sabah son adım."},
  {id:81,barcode:"8690104050307",brand:"Dermoskin",name:"Acne Spot Treatment Krem",category:"tedavi",usage:"pm",basePrice:150,actives:["Salisilik Asit","Çinko","Niasinamid"],pairs_with:[77],certs:["dermo-test"],boycott:"temiz",pregnancy_safe:"consult",skin_types:["oily"],concerns:["acne"],desc:"Sivilce nokta tedavisi için krem.",how_to:"Akşam sivilce üzerine nokta uygulayın."},
  {id:82,barcode:"8690104050314",brand:"Dermoskin",name:"Sebum Control Tonik",category:"tonik",usage:"both",basePrice:175,actives:["Çinko","Niasinamid","Panthenol"],pairs_with:[81,80],certs:["dermo-test"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["oily"],concerns:["acne"],desc:"Yağ kontrolü için günlük tonik.",how_to:"Temizleyici sonrası pamukla silin."},
  {id:83,barcode:"8690104050321",brand:"Dermoskin",name:"Hydra Intense Serum",category:"serum",usage:"both",basePrice:220,actives:["Hyalüronik Asit","Panthenol"],pairs_with:[79,80],certs:["dermo-test","fragrance-free"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["dry","normal","sensitive"],concerns:["moisture"],desc:"Yoğun nem için hyalüronik asit serumu.",how_to:"Temizleyici sonrası 2–3 damla."},
  {id:84,barcode:"8690104050338",brand:"Dermoskin",name:"Age Defense Gece Kremi",category:"nemlendirici",usage:"pm",basePrice:245,actives:["Retinol","Peptid","Vitamin E"],pairs_with:[83],certs:["dermo-test"],boycott:"temiz",pregnancy_safe:"avoid",skin_types:["normal","dry"],concerns:["antiage"],desc:"Retinol içerikli yaşlanma karşıtı gece bakımı.",how_to:"Akşam temiz cilde uygulayın."},
  {id:85,barcode:"8690104050345",brand:"Dermoskin",name:"Micellar Temizleme Suyu",category:"temizleyici",usage:"both",basePrice:165,actives:["Panthenol","Aloe Vera"],pairs_with:[83,79],certs:["dermo-test","fragrance-free"],boycott:"temiz",pregnancy_safe:"safe",skin_types:["normal","sensitive","dry"],concerns:["moisture"],desc:"Nazik günlük makyaj temizleyici micellar su.",how_to:"Pamukla sildikten sonra durulamayın."},
];

// ══════════════════════════════════════════════════════════════
// ANAHTAR/TAHMIN DEĞERLERİ
// ══════════════════════════════════════════════════════════════
function getUrlMode(){
  try{
    const p=new URLSearchParams(window.location.search);
    if(p.has("superadmin"))return"superadmin";
    if(p.has("pharmacy"))return"pharmacy";
    if(p.has("admin"))return"admin";
  }catch{}
  return"customer";
}
// Kiosk linkindeki ?store=<eczane_id> — giriş yapılmadan, hangi eczanenin
// kataloğunun gösterileceğini belirler. Fiziksel/sabit kiosk cihazları için.
function getStoreParam(){
  try{
    const p=new URLSearchParams(window.location.search);
    return p.get("store")||null;
  }catch{return null;}
}

// ══════════════════════════════════════════════════════════════
// ANKET
// ══════════════════════════════════════════════════════════════
const QUIZ_STEPS = [
  {
    id:"skin_type",question:"Cildiniz genellikle nasıl?",emoji:"🧴",
    options:[
      {id:"oily",    label:"Yağlı / Karma",    emoji:"💧",desc:"T-bölgesi parlıyor, gözenekler belirgin"},
      {id:"dry",     label:"Kuru",              emoji:"🌵",desc:"Gerginlik hissediyorum, soyulma olabiliyor"},
      {id:"sensitive",label:"Hassas / Reaktif", emoji:"🌸",desc:"Kolayca kızarıyor, tahriş oluyor"},
      {id:"normal",  label:"Normal / Dengeli",  emoji:"✨",desc:"Ne çok yağlı ne çok kuru"},
    ]
  },
  {
    id:"concern",question:"En çok önem verdiğiniz nedir?",emoji:"🎯",
    options:[
      {id:"moisture",label:"Nem & Yatıştırma",       emoji:"💦",desc:"Cildim susuz, beslenmeli"},
      {id:"acne",    label:"Akne & Gözenek",         emoji:"🫧",desc:"Sivilce ve gözenek sorunum var"},
      {id:"pigment", label:"Leke & Ton Eşitsizliği", emoji:"🌟",desc:"Lekeler ve mat görünüm"},
      {id:"antiage", label:"Yaşlanma Karşıtı",       emoji:"⏳",desc:"Kırışık ve sarkma önlemi"},
      {id:"sun",     label:"Güneş Koruması",          emoji:"☀️",desc:"Güneş hasarı ve SPF"},
    ]
  },
  {
    id:"pregnant",question:"Hamile veya emziriyor musunuz?",emoji:"🤰",
    options:[
      {id:"no",           label:"Hayır",                  emoji:"🙅",desc:""},
      {id:"pregnant",     label:"Evet, hamileyim",        emoji:"🤰",desc:"Bazı aktiflerden kaçınmanız gerekebilir"},
      {id:"breastfeeding",label:"Emziriyorum",            emoji:"👶",desc:"Bazı ürünler önerilmeyebilir"},
      {id:"unsure",       label:"Henüz bilmiyorum",       emoji:"🤔",desc:"İhtiyatlı ürünleri göster"},
    ]
  },
  {
    id:"actives",question:"Güçlü aktif kullanıyor musunuz?",emoji:"🔬",
    options:[
      {id:"yes", label:"Evet, düzenli kullanıyorum",  emoji:"💪",desc:"Retinol, AHA, BHA gibi ürünler"},
      {id:"new", label:"Yeni başlıyorum",              emoji:"🌱",desc:"Denemek istiyorum ama bilmiyorum"},
      {id:"no",  label:"Hayır, hassas kalıyorum",     emoji:"🛡️",desc:"Minimal, yatıştırıcı tercih ederim"},
    ]
  },
];

// ══════════════════════════════════════════════════════════════
// ÖNERI ALGORİTMASI (AI yok — tamamen kural tabanlı)
// ══════════════════════════════════════════════════════════════
function scoreProduct(product, profile, isRecommendation=false) {
  let score = 0;
  if (!profile) return score;
  const {skin_type, concern, pregnant, actives_level} = profile;
  const pSkinTypes = product.skin_types || [];
  const pConcerns = product.concerns || [];

  // 1. Hamilelik güvenliği (en yüksek öncelik)
  const isPregnantUser = pregnant === "pregnant" || pregnant === "breastfeeding" || pregnant === "unsure";
  if (isPregnantUser) {
    if (product.pregnancy_safe === "avoid") return -99; // listeye alma
    if (product.pregnancy_safe === "safe") score += 3;
    if (product.pregnancy_safe === "consult") score += 0; // nötr
  }

  // 2. Cilt tipi uyumu
  if (pSkinTypes.includes(skin_type)) score += 3;

  // 3. Endişe uyumu
  if (pConcerns.includes(concern)) score += 3;

  // 4. Aktif deneyimi
  const hasRetinol = (product.actives||[]).some(a=>a.toLowerCase().includes("retinol"));
  const hasStrongAcid = (product.actives||[]).some(a=>a.toLowerCase().includes("salisilik")||a.toLowerCase().includes("glikolik"));
  if (actives_level === "no" && (hasRetinol || hasStrongAcid)) score -= 2;
  if (actives_level === "new" && (hasRetinol || hasStrongAcid)) score += 1; // rehberlik et

  // Öneri modunda ek ağırlıklar
  if (isRecommendation) {
    // 5. Kâr marjı skoru
    const price = product.price || product.basePrice;
    const cost = product.cost;
    if (cost && cost > 0) {
      const margin = (price - cost) / price;
      score += Math.round(margin * 4); // max +4
    } else {
      score += 1; // maliyet bilinmiyorsa hafif pozitif
    }

    // 6. SKT aciliyeti
    if (product.skt) {
      const daysLeft = Math.round((new Date(product.skt) - new Date()) / (1000*60*60*24));
      if (daysLeft < 30) score += 5;
      else if (daysLeft < 60) score += 3;
      else if (daysLeft < 90) score += 1;
    }
  }

  return Math.max(0, score);
}

function getAlgoRecs(basket, catalog, profile) {
  const bIds = new Set(basket.map(p=>p.id));

  return catalog
    .filter(p => hasStock(p) && !bIds.has(p.id))
    .map(p => {
      const pairsScore = basket.filter(b=>(b.pairs_with||[]).includes(p.id)).length * 3;
      const profScore = scoreProduct(p, profile, true);
      if (profScore < 0) return null; // hamile için tehlikeli ürünleri çıkar
      const total = pairsScore + profScore;
      const reason = buildReason(p, profile, pairsScore > 0, basket);
      return {prod: p, score: total, reason};
    })
    .filter(Boolean)
    .filter(r => r.score > 0)
    .sort((a,b) => b.score - a.score)
    .slice(0, 4);
}

function buildReason(product, profile, isPair, basket) {
  if (!profile) {
    if (isPair) return `${basket.find(b=>(b.pairs_with||[]).includes(product.id))?.name?.split(" ")[0]} ile uyumlu`;
    return "Kataloğunuzdan öneri";
  }
  const {skin_type, concern, pregnant} = profile;
  const isPregnant = pregnant === "pregnant" || pregnant === "breastfeeding";
  if (isPregnant && product.pregnancy_safe === "safe") return "Hamilelikte güvenli ✓";
  if (isPair) {
    const pairName = basket.find(b=>(b.pairs_with||[]).includes(product.id))?.name?.split(" ").slice(0,2).join(" ");
    return `${pairName} ile tamamlayıcı`;
  }
  if ((product.skin_types||[]).includes(skin_type)) return `${SKIN_TYPE_LABELS[skin_type]} cilt için`;
  if ((product.concerns||[]).includes(concern)) return `${CONCERN_LABELS[concern]} için`;
  if (product.cost) {
    const m = Math.round(((product.price||product.basePrice) - product.cost) / (product.price||product.basePrice) * 100);
    if (m > 30) return "Yüksek kârlı ürün";
  }
  if (product.skt) {
    const days = Math.round((new Date(product.skt)-new Date())/(1000*60*60*24));
    if (days < 60) return `SKT yakın — ${days} gün`;
  }
  return "Profilinizle uyumlu";
}

// ══════════════════════════════════════════════════════════════
// STOK YARDIMCILARI (stock: number=adet, true=sınırsız/bilinmiyor, false=yok)
// ══════════════════════════════════════════════════════════════
function hasStock(p){
  if(p==null)return false;
  if(p.stock===undefined||p.stock===null)return true;
  if(typeof p.stock==="boolean")return p.stock;
  if(typeof p.stock==="number")return p.stock>0;
  return true;
}
function stockQty(p){
  return typeof p?.stock==="number"?p.stock:null;
}
function daysUntil(dateStr){
  if(!dateStr)return null;
  return Math.ceil((new Date(dateStr)-new Date())/(1000*60*60*24));
}

// ══════════════════════════════════════════════════════════════
// ÜRÜN HAREKETLİLİĞİ ETİKETİ (satış hızına göre)
// ══════════════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════════════
// SATIŞ HIZINI GERÇEK VERİDEN OTOMATİK HESAPLA
// Son 28 günün gerçek satışlarından haftalık ortalama çıkarır.
// Yeterli veri yoksa (yeni ürün/eczane) elle girilen değere düşer.
// ══════════════════════════════════════════════════════════════
function computeSalesVelocity(itemId,sales){
  if(!sales||sales.length===0)return null;
  const now=new Date();
  const since=new Date(now);since.setDate(now.getDate()-28);
  let qty=0;
  sales.forEach(s=>{
    const d=new Date(s.date);
    if(d<since)return;
    (s.items||[]).forEach(it=>{if(it.id===itemId)qty+=(it.qty||1);});
  });
  if(qty===0)return null;
  const daysSpan=Math.max(1,Math.round((now-since)/86400000));
  return Math.round((qty/daysSpan)*7*10)/10;
}
function effectiveWeeklySales(item,sales){
  const computed=computeSalesVelocity(item.id,sales);
  return computed!=null?computed:(item.weeklySales??null);
}

function getMovement(item,sales){
  const eff=effectiveWeeklySales(item,sales);
  const w=Number(eff);
  if(eff==null||isNaN(w))
    return{label:"Bilinmiyor",color:"#888",bg:"#F4F4F4"};
  if(w<=0)return{label:"Hareketsiz",color:"#8B2E2E",bg:"#FDECEA"};
  if(w<2)return{label:"Yavaş",color:"#7D4700",bg:"#FFF3E0"};
  if(w<6)return{label:"Normal",color:"#1E6B3E",bg:"#E9F7EF"};
  return{label:"Hızlı",color:"#1A5276",bg:"#D6EAF8"};
}

// ══════════════════════════════════════════════════════════════
// OTO İNDİRİM ÖNERİSİ — stok + satış hızı + miad + fiyat + kârlılık
// Not: Kural tabanlı bir yardımcı gösterge; kesin bir mali hesap değildir.
// ══════════════════════════════════════════════════════════════
function suggestDiscount(item,sales){
  const price=item.discountedPrice||item.price||item.basePrice;
  const cost=item.cost;
  const qty=stockQty(item);
  const weekly=Number(effectiveWeeklySales(item,sales))||0;
  const days=daysUntil(item.skt);

  let score=0;
  const reasons=[];

  // 1) Stok devri (days-of-supply = stok / günlük satış)
  if(qty!=null&&qty>0){
    if(weekly>0){
      const dos=qty/(weekly/7);
      if(dos>120){score+=40;reasons.push(`Stok, satış hızına göre ~${Math.round(dos)} günlük — çok yüksek`);}
      else if(dos>60){score+=22;reasons.push(`Stok devri yavaş (~${Math.round(dos)} günlük)`);}
      else if(dos>35){score+=10;reasons.push("Stok devri ortalamanın altında");}
    } else {
      score+=20;reasons.push("Son 28 günde satış yok — stok bekliyor");
    }
  }

  // 2) Miad aciliyeti (SKT)
  if(days!=null&&days>=0){
    if(days<=30){score+=45;reasons.push(`SKT'ye ${days} gün kaldı`);}
    else if(days<=60){score+=28;reasons.push(`SKT'ye ${days} gün kaldı`);}
    else if(days<=90){score+=14;reasons.push(`SKT'ye ${days} gün kaldı`);}
  }

  if(score<=0||reasons.length===0)return null;

  // 3) Fiyat + kârlılık sınırı — MIN_MARGIN altına düşürme
  let percent=Math.min(35,Math.round(score/2.2));
  if(cost&&cost>0&&price>0){
    const minPriceForMargin=cost/(1-MIN_MARGIN);
    const maxPct=Math.floor((1-(minPriceForMargin/price))*100);
    percent=Math.min(percent,Math.max(0,maxPct));
  }
  if(percent<5)return null;

  const urgency=score>=55?"high":score>=25?"medium":"low";
  return{percent,reasons,urgency};
}

// ══════════════════════════════════════════════════════════════
// BARKOD GÖRSELİ (dekoratif — gerçek EAN-13 kodlaması değildir,
// sayı her zaman altında okunur şekilde yazılır)
// ══════════════════════════════════════════════════════════════
function Barcode({value,height=30}){
  const digits=String(value||"").replace(/\D/g,"");
  const src=digits||"0000000000000";
  const bars=[];
  let x=2;
  for(let i=0;i<src.length;i++){
    const d=parseInt(src[i],10);
    const w=1+(d%4);
    if(i%2===0)bars.push(<rect key={i} x={x} y={0} width={w} height={height} fill="#1C1C1A"/>);
    x+=w+1.4;
  }
  const totalWidth=x+2;
  if(!digits)return null;
  return(
    <div style={{textAlign:"center"}}>
      <svg viewBox={`0 0 ${totalWidth} ${height}`} width="100%" height={height} style={{display:"block",maxWidth:180,margin:"0 auto"}}>
        <rect x={0} y={0} width={totalWidth} height={height} fill="#fff"/>
        {bars}
      </svg>
      <div style={{fontSize:9,letterSpacing:1.5,color:"#888",fontFamily:"monospace",marginTop:2}}>{digits}</div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════
// KÜÇÜK BİLEŞENLER
// ══════════════════════════════════════════════════════════════
function useTheme(s){return BRAND_THEMES[s?.primaryBrand]||DT;}

// ══════════════════════════════════════════════════════════════
// SUPABASE SATIR <-> JS NESNE DÖNÜŞÜMLERİ
// ══════════════════════════════════════════════════════════════
const poolRowToJs=(r)=>({
  id:r.id, barcode:r.barcode||"", brand:r.brand, name:r.name,
  category:r.category, usage:r.usage, basePrice:r.base_price,
  actives:r.actives||[], pairs_with:r.pairs_with||[], certs:r.certs||[],
  skin_types:r.skin_types||[], concerns:r.concerns||[],
  boycott:r.boycott, pregnancy_safe:r.pregnancy_safe,
  photo:r.photo||null, desc:r.description||"", how_to:r.how_to||"",
});
const POOL_FIELD_MAP={basePrice:"base_price",desc:"description"};

const catalogRowToJs=(r)=>({
  ...(r.pool?poolRowToJs(r.pool):{id:r.product_id}),
  price:r.price, discountedPrice:r.discounted_price,
  stock:r.stock==null?true:r.stock, cost:r.cost, weeklySales:r.weekly_sales, skt:r.skt,
});

// ══════════════════════════════════════════════════════════════
// RESPONSIVE CSS (PC / Tablet uyumu) — bir kez <head>'e enjekte edilir
// ══════════════════════════════════════════════════════════════
function useResponsiveStyles(){
  useEffect(()=>{
    if(document.getElementById("eczane-responsive-style"))return;
    const style=document.createElement("style");
    style.id="eczane-responsive-style";
    style.textContent=`
      .kiosk-shell{max-width:430px;margin:0 auto;}
      @media (min-width:640px){ .kiosk-shell{max-width:480px;} }

      .admin-shell{max-width:430px;margin:0 auto;}
      @media (min-width:760px){ .admin-shell{max-width:840px;} }
      @media (min-width:1100px){ .admin-shell{max-width:1180px;} }

      .admin-grid{display:flex;flex-direction:column;gap:8px;}
      @media (min-width:760px){
        .admin-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));
          align-items:start;gap:10px;}
      }

      /* Genel geçiş & etkileşim iyileştirmeleri */
      *{box-sizing:border-box;}
      button{transition:transform .1s ease,opacity .15s ease,filter .15s ease;
        -webkit-tap-highlight-color:transparent;}
      button:active:not(:disabled){transform:scale(.97);}
      button:disabled{cursor:default;}
      button:not(:disabled):hover{filter:brightness(1.04);}
      input,select,textarea{transition:border-color .15s ease,box-shadow .15s ease;}
      input:focus,select:focus,textarea:focus{outline:none;border-color:#1C1C1A !important;
        box-shadow:0 0 0 3px rgba(28,28,26,.08);}

      /* Sekme/ekran geçişlerinde yumuşak beliriş */
      .fade-in{animation:eczFadeIn .22s ease both;}
      @keyframes eczFadeIn{from{opacity:0;transform:translateY(4px);}to{opacity:1;transform:translateY(0);}}

      /* Dönen yükleniyor göstergesi */
      .ecz-spinner{width:22px;height:22px;border-radius:50%;
        border:2.5px solid rgba(0,0,0,.12);border-top-color:#1C1C1A;
        animation:eczSpin .7s linear infinite;display:inline-block;}
      .ecz-spinner.light{border:2.5px solid rgba(255,255,255,.25);border-top-color:#fff;}
      @keyframes eczSpin{to{transform:rotate(360deg);}}

      /* İskelet yükleniyor kartı */
      .ecz-skeleton{background:linear-gradient(90deg,#EFEFEF 25%,#F7F7F7 37%,#EFEFEF 63%);
        background-size:400% 100%;animation:eczShimmer 1.4s ease infinite;border-radius:10px;}
      @keyframes eczShimmer{0%{background-position:100% 50%;}100%{background-position:0 50%;}}
    `;
    document.head.appendChild(style);
  },[]);
}
// ══════════════════════════════════════════════════════════════
// BOŞTA KALMA ZAMAN AŞIMI — belirli süre etkileşim olmazsa tetiklenir
// (kiosk'ta unutulan sepeti/oturumu güvenli şekilde sıfırlamak için)
// ══════════════════════════════════════════════════════════════
function useIdleTimer(timeoutMs,onIdle,enabled=true){
  useEffect(()=>{
    if(!enabled)return;
    let timer;
    const reset=()=>{clearTimeout(timer);timer=setTimeout(onIdle,timeoutMs);};
    const events=["mousedown","mousemove","keydown","touchstart","scroll","click"];
    events.forEach(e=>window.addEventListener(e,reset,{passive:true}));
    reset();
    return()=>{clearTimeout(timer);events.forEach(e=>window.removeEventListener(e,reset));};
  },[timeoutMs,onIdle,enabled]);
}
function Spinner({light=false,size=22}){
  return <span className={`ecz-spinner${light?" light":""}`} style={{width:size,height:size}}/>;
}
function SkeletonCard(){
  return(
    <div style={{background:"#fff",borderRadius:14,border:"1px solid #E0E0E0",padding:14,marginBottom:10}}>
      <div style={{display:"flex",gap:10,alignItems:"center"}}>
        <div className="ecz-skeleton" style={{width:40,height:40,borderRadius:8,flexShrink:0}}/>
        <div style={{flex:1}}>
          <div className="ecz-skeleton" style={{width:"40%",height:9,borderRadius:5,marginBottom:6}}/>
          <div className="ecz-skeleton" style={{width:"70%",height:12,borderRadius:5,marginBottom:6}}/>
          <div className="ecz-skeleton" style={{width:"30%",height:11,borderRadius:5}}/>
        </div>
      </div>
    </div>
  );
}
function Avatar({brand,size=42}){
  const s=BRAND_COLORS[brand]||{bg:"#F0F0F0",c:"#333"};
  return <div style={{width:size,height:size,borderRadius:Math.round(size*.25),background:s.bg,color:s.c,
    display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:size*.28,flexShrink:0}}>
    {brand.slice(0,2).toUpperCase()}</div>;
}
function BoycottBadge({status}){
  const m=BOYCOTT_META[status]||BOYCOTT_META.bilinmiyor;
  return <span style={{background:m.bg,color:m.color,fontSize:10,fontWeight:700,padding:"2px 8px",borderRadius:5}}>{m.label}</span>;
}
function PregBadge({status}){
  const m=PREG_META[status||"unknown"];if(!m||status==="unknown")return null;
  return <span style={{background:m.bg,color:m.color,fontSize:10,fontWeight:600,padding:"2px 8px",borderRadius:5}}>{m.label}</span>;
}
function CertBadge({id}){
  const m=CERT_META[id];if(!m)return null;
  return <span style={{background:m.bg,color:m.color,fontSize:10,fontWeight:500,padding:"2px 7px",borderRadius:5}}>{m.label}</span>;
}
function Sw({value,onChange,color="#2C4A3E"}){
  return <div onClick={()=>onChange(!value)} style={{width:44,height:24,borderRadius:12,
    background:value?color:"#DDD",cursor:"pointer",position:"relative",transition:"background .2s"}}>
    <div style={{position:"absolute",top:3,left:value?22:2,width:18,height:18,borderRadius:9,
      background:"#fff",transition:"left .2s",boxShadow:"0 1px 3px rgba(0,0,0,.2)"}}/>
  </div>;
}
function Toast({msg}){
  if(!msg)return null;
  return <div style={{background:"#E9F7EF",color:"#1E6B3E",borderRadius:8,padding:"9px 13px",
    fontSize:12,fontWeight:600,marginBottom:10,border:"1px solid #C3E6CB"}}>{msg}</div>;
}

// ══════════════════════════════════════════════════════════════
// CİLT ANKETİ
// ══════════════════════════════════════════════════════════════
function SkinQuiz({onComplete,theme}){
  const T=theme||DT;
  const [step,setStep]=useState(0);
  const [answers,setAnswers]=useState({});
  const current=QUIZ_STEPS[step];
  const progress=(step/QUIZ_STEPS.length)*100;

  const select=(optId)=>{
    const na={...answers,[current.id]:optId};
    setAnswers(na);
    if(step<QUIZ_STEPS.length-1){
      setTimeout(()=>setStep(s=>s+1),200);
    } else {
      setTimeout(()=>onComplete({
        skin_type:na.skin_type,concern:na.concern,
        pregnant:na.pregnant,actives_level:na.actives,
      }),200);
    }
  };

  return(
    <div style={{position:"fixed",inset:0,zIndex:400,
      background:"rgba(0,0,0,.6)",backdropFilter:"blur(4px)",
      display:"flex",alignItems:"flex-end",justifyContent:"center",
      fontFamily:"'Inter',sans-serif"}}>
      <div className="kiosk-shell" style={{background:"#fff",borderRadius:"24px 24px 0 0",width:"100%",
        padding:"24px 20px 48px",maxHeight:"92vh",overflowY:"auto"}}>
        {/* Progress */}
        <div style={{height:3,background:"#F0F0F0",borderRadius:2,marginBottom:20,overflow:"hidden"}}>
          <div style={{height:3,background:T.primary,borderRadius:2,
            width:`${progress}%`,transition:"width .4s ease"}}/>
        </div>
        <div style={{fontSize:11,color:"#AAA",fontWeight:500,marginBottom:4,letterSpacing:.5}}>
          {step+1} / {QUIZ_STEPS.length}
        </div>
        <div style={{fontSize:30,marginBottom:6}}>{current.emoji}</div>
        <div style={{fontSize:19,fontWeight:700,color:"#1C1C1A",lineHeight:1.25,marginBottom:18}}>
          {current.question}
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:9}}>
          {current.options.map(opt=>{
            const sel=answers[current.id]===opt.id;
            return(
              <button key={opt.id} onClick={()=>select(opt.id)}
                style={{display:"flex",alignItems:"center",gap:12,padding:"13px 14px",
                  borderRadius:13,textAlign:"left",cursor:"pointer",
                  border:`2px solid ${sel?T.primary:"#EBEBEB"}`,
                  background:sel?T.bg:"#FAFAFA",transition:"all .15s"}}>
                <span style={{fontSize:22,flexShrink:0}}>{opt.emoji}</span>
                <div style={{flex:1}}>
                  <div style={{fontSize:14,fontWeight:600,color:sel?T.primary:"#1C1C1A"}}>{opt.label}</div>
                  {opt.desc&&<div style={{fontSize:11,color:"#888",marginTop:1,lineHeight:1.4}}>{opt.desc}</div>}
                </div>
                {sel&&<div style={{marginLeft:"auto",fontSize:16,color:T.primary,flexShrink:0}}>✓</div>}
              </button>
            );
          })}
        </div>
        <button onClick={()=>onComplete(null)}
          style={{width:"100%",marginTop:16,background:"none",border:"none",
            color:"#BBB",fontSize:12,cursor:"pointer",padding:"8px 0"}}>
          Anketi atla
        </button>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════
// BARKOD TARAYICI
// mode="customer"  → ürün ara, sepete ekle
// mode="pharmacy"  → ürün ara, kataloğa ekle
// mode="fill"      → sadece barkod değerini döndür (form doldurma)
// ══════════════════════════════════════════════════════════════
function BarcodeScanner({catalog=[], onFound, onFill, onClose, mode="customer"}){
  const videoRef  = useRef(null);
  const streamRef = useRef(null);
  const timerRef  = useRef(null);
  const detRef    = useRef(null);

  // status: "init" | "cam_error" | "no_api" | "scanning" | "found" | "notfound" | "manual"
  const [status, setStatus]   = useState("init");
  const [manual, setManual]   = useState("");
  const [manualErr, setManualErr] = useState("");
  const [found,  setFound]    = useState(null);
  const [camErr, setCamErr]   = useState("");

  // Temizleme
  const cleanup = () => {
    if(timerRef.current){ clearInterval(timerRef.current); timerRef.current=null; }
    if(streamRef.current){ streamRef.current.getTracks().forEach(t=>t.stop()); streamRef.current=null; }
  };

  useEffect(()=>{
    // BarcodeDetector API desteği var mı?
    if(!("BarcodeDetector" in window)){
      setStatus("no_api");
      return;
    }
    // Desteklenen formatları sorgula
    window.BarcodeDetector.getSupportedFormats?.().then(fmts=>{
      const wanted=["ean_13","ean_8","code_128","upc_a","code_39"];
      const use=fmts.filter(f=>wanted.includes(f));
      detRef.current = new window.BarcodeDetector({formats: use.length ? use : wanted});
    }).catch(()=>{
      detRef.current = new window.BarcodeDetector({formats:["ean_13","ean_8","code_128","upc_a"]});
    });
    startCamera();
    return cleanup;
  }, []);

  const startCamera = async () => {
    setStatus("init");
    try {
      // Önce arka kamera, yoksa herhangi kamera
      let stream;
      try {
        stream = await navigator.mediaDevices.getUserMedia({
          video:{ facingMode:{ideal:"environment"}, width:{ideal:1280}, height:{ideal:720} }
        });
      } catch {
        stream = await navigator.mediaDevices.getUserMedia({ video:true });
      }
      streamRef.current = stream;

      // Video elemanı henüz render edilmemiş olabilir — kısa bekle
      await new Promise(r => setTimeout(r, 100));

      if(!videoRef.current){ cleanup(); setStatus("cam_error"); setCamErr("Video başlatılamadı."); return; }
      videoRef.current.srcObject = stream;

      // play() Promise'ini yakala
      try { await videoRef.current.play(); } catch(e){ /* autoplay policy — ilerle */ }

      setStatus("scanning");

      // setInterval — iframe'de requestAnimationFrame güvenilmez
      timerRef.current = setInterval(async () => {
        if(!videoRef.current || !detRef.current) return;
        if(videoRef.current.readyState < 2) return; // HAVE_CURRENT_DATA
        try {
          const barcodes = await detRef.current.detect(videoRef.current);
          if(barcodes.length > 0){
            clearInterval(timerRef.current);
            timerRef.current = null;
            processCode(barcodes[0].rawValue);
          }
        } catch { /* devam */ }
      }, 400); // 400ms aralıkla tara

    } catch(e) {
      const msg = e.name === "NotAllowedError"
        ? "Kamera izni reddedildi. Tarayıcı ayarlarından izin verin."
        : e.name === "NotFoundError"
        ? "Kamera bulunamadı."
        : `Kamera açılamadı: ${e.message}`;
      setCamErr(msg);
      setStatus("cam_error");
    }
  };

  const processCode = (raw) => {
    const code = raw.trim();
    if(mode === "fill"){
      // Sadece barkod değerini döndür — ürün arama yapma
      cleanup();
      if(onFill) onFill(code);
      onClose();
      return;
    }
    const product = catalog.find(p =>
      p.barcode === code ||
      p.barcode === code.replace(/^0+/,"") ||
      ("0"+p.barcode) === code
    );
    if(product){ setFound(product); setStatus("found"); }
    else { setStatus("notfound"); }
  };

  const retryCamera = () => {
    cleanup();
    setFound(null);
    setCamErr("");
    if(!("BarcodeDetector" in window)){ setStatus("no_api"); return; }
    startCamera();
  };

  const handleManualSearch = () => {
    const code = manual.trim();
    if(!code){ setManualErr("Barkod numarası girin."); return; }
    if(mode==="fill"){
      cleanup();
      if(onFill) onFill(code);
      onClose();
      return;
    }
    setManualErr("");
    processCode(code);
  };

  const modeLabel = mode==="fill"
    ? "Barkod değerini oku"
    : mode==="pharmacy"
    ? "Havuzdan kataloğa ekle"
    : "Ürünü sepete ekle";

  return(
    <div style={{position:"fixed",inset:0,zIndex:300,background:"rgba(0,0,0,.92)",
      display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",
      fontFamily:"'Inter',sans-serif",padding:16}}>

      <button onClick={()=>{ cleanup(); onClose(); }}
        style={{position:"absolute",top:16,right:16,background:"rgba(255,255,255,.15)",border:"none",
          color:"#fff",borderRadius:8,padding:"8px 14px",fontSize:13,cursor:"pointer",zIndex:10}}>
        ✕ Kapat
      </button>

      <div style={{color:"#fff",fontSize:17,fontWeight:700,marginBottom:12,textAlign:"center"}}>
        📷 Barkod Tara
        <div style={{fontSize:11,opacity:.55,fontWeight:400,marginTop:3}}>{modeLabel}</div>
      </div>

      {/* Kamera görüntüsü — her zaman render'da, sadece status'a göre görünür */}
      <div style={{
        position:"relative",width:280,height:220,borderRadius:14,overflow:"hidden",
        border:"2px solid rgba(255,255,255,.25)",marginBottom:12,
        display: status==="scanning" ? "block" : "none"
      }}>
        <video
          ref={videoRef}
          style={{width:"100%",height:"100%",objectFit:"cover",display:"block"}}
          playsInline muted autoPlay
        />
        {/* Tarama çizgisi */}
        <div style={{position:"absolute",inset:0,pointerEvents:"none"}}>
          <div style={{position:"absolute",top:"50%",left:"10%",right:"10%",height:2,
            background:"rgba(74,207,130,.9)",boxShadow:"0 0 6px rgba(74,207,130,.8)",
            animation:"scan 1.5s ease-in-out infinite"}}/>
          {/* Köşe kılavuzları */}
          {[[0,0,"borderTop borderLeft"],[0,1,"borderTop borderRight"],
            [1,0,"borderBottom borderLeft"],[1,1,"borderBottom borderRight"]].map(([r,c,_],i)=>(
            <div key={i} style={{position:"absolute",
              top:r?"auto":8,bottom:r?8:"auto",
              left:c?"auto":8,right:c?8:"auto",
              width:20,height:20,
              borderTop:!r?"2px solid rgba(255,255,255,.6)":"none",
              borderBottom:r?"2px solid rgba(255,255,255,.6)":"none",
              borderLeft:!c?"2px solid rgba(255,255,255,.6)":"none",
              borderRight:c?"2px solid rgba(255,255,255,.6)":"none"}}/>
          ))}
        </div>
      </div>

      {/* Durum: başlıyor */}
      {status==="init"&&(
        <div style={{color:"rgba(255,255,255,.7)",fontSize:13,textAlign:"center",marginBottom:12}}>
          ⏳ Kamera başlatılıyor…
        </div>
      )}

      {/* Durum: kamera hatası */}
      {(status==="cam_error"||status==="no_api")&&(
        <div style={{background:"#fff",borderRadius:14,padding:18,maxWidth:300,width:"100%",textAlign:"center",marginBottom:12}}>
          <div style={{fontSize:24,marginBottom:8}}>📵</div>
          <div style={{fontSize:13,fontWeight:600,color:"#1C1C1A",marginBottom:4}}>
            {status==="no_api"?"Kamera tarama desteklenmiyor":"Kamera açılamadı"}
          </div>
          <div style={{fontSize:11,color:"#888",marginBottom:12,lineHeight:1.5}}>
            {status==="no_api"
              ?"Bu tarayıcı barkod okuma API'sini desteklemiyor (Chrome 88+ gerekli)."
              :camErr}
          </div>
          {status==="cam_error"&&(
            <button onClick={retryCamera}
              style={{width:"100%",background:"#2C4A3E",color:"#fff",border:"none",borderRadius:9,
                padding:"9px",fontSize:13,fontWeight:600,cursor:"pointer",marginBottom:8}}>
              Tekrar Dene
            </button>
          )}
          <button onClick={()=>setStatus("manual")}
            style={{width:"100%",background:"#F0F0F0",color:"#333",border:"none",borderRadius:9,
              padding:"9px",fontSize:13,cursor:"pointer"}}>
            Manuel Barkod Girişi
          </button>
        </div>
      )}

      {/* Durum: taranıyor */}
      {status==="scanning"&&(
        <div style={{textAlign:"center"}}>
          <div style={{color:"rgba(255,255,255,.6)",fontSize:12,marginBottom:8}}>
            Barkodu çerçeve içine getirin
          </div>
          <button onClick={()=>{ cleanup(); setStatus("manual"); }}
            style={{background:"none",border:"1px solid rgba(255,255,255,.3)",
              color:"rgba(255,255,255,.65)",borderRadius:8,padding:"5px 14px",fontSize:11,cursor:"pointer"}}>
            Manuel giriş
          </button>
        </div>
      )}

      {/* Durum: ürün bulundu */}
      {status==="found"&&found&&(
        <div style={{background:"#fff",borderRadius:16,padding:20,maxWidth:300,width:"100%",textAlign:"center"}}>
          <div style={{fontSize:28,marginBottom:6}}>✅</div>
          <div style={{fontSize:11,color:"#888",marginBottom:1}}>{found.brand}</div>
          <div style={{fontSize:16,fontWeight:700,color:"#1C1C1A",marginBottom:4,lineHeight:1.3}}>{found.name}</div>
          <div style={{fontSize:11,color:"#AAA",marginBottom:8}}>{found.barcode}</div>
          <div style={{fontSize:20,fontWeight:700,color:"#2C4A3E",marginBottom:16}}>
            {(found.price||found.basePrice).toLocaleString("tr-TR")} ₺
          </div>
          <div style={{display:"flex",gap:8}}>
            <button onClick={()=>{ if(onFound) onFound(found); cleanup(); onClose(); }}
              style={{flex:1,background:"#2C4A3E",color:"#fff",border:"none",borderRadius:10,
                padding:"11px",fontSize:13,fontWeight:700,cursor:"pointer"}}>
              {mode==="customer"?"Sepete Ekle ✓":"Kataloğa Ekle ✓"}
            </button>
            <button onClick={retryCamera}
              style={{background:"#F0F0F0",color:"#333",border:"none",borderRadius:10,
                padding:"11px 12px",fontSize:12,cursor:"pointer"}}>Yeni Tara</button>
          </div>
        </div>
      )}

      {/* Durum: ürün bulunamadı */}
      {status==="notfound"&&(
        <div style={{background:"#fff",borderRadius:16,padding:20,maxWidth:300,width:"100%",textAlign:"center"}}>
          <div style={{fontSize:28,marginBottom:8}}>❓</div>
          <div style={{fontSize:15,fontWeight:600,color:"#1C1C1A",marginBottom:4}}>Ürün bulunamadı</div>
          <div style={{fontSize:12,color:"#888",marginBottom:14,lineHeight:1.5}}>
            Bu barkod sistemde kayıtlı değil.
          </div>
          <button onClick={retryCamera}
            style={{width:"100%",background:"#2C4A3E",color:"#fff",border:"none",borderRadius:10,
              padding:"10px",fontSize:13,fontWeight:700,cursor:"pointer",marginBottom:8}}>
            Tekrar Tara
          </button>
          <button onClick={()=>setStatus("manual")}
            style={{width:"100%",background:"#F0F0F0",color:"#333",border:"none",borderRadius:10,
              padding:"10px",fontSize:13,cursor:"pointer"}}>
            Manuel Giriş
          </button>
        </div>
      )}

      {/* Durum: manuel giriş */}
      {status==="manual"&&(
        <div style={{background:"#fff",borderRadius:16,padding:20,maxWidth:300,width:"100%"}}>
          <div style={{fontSize:14,fontWeight:700,color:"#1C1C1A",marginBottom:4}}>
            {mode==="fill"?"Barkod Numarasını Girin":"Barkod ile Ara"}
          </div>
          <div style={{fontSize:11,color:"#888",marginBottom:12,lineHeight:1.4}}>
            {mode==="fill"
              ?"Numara girilince forma otomatik aktarılır."
              :"Kutu ya da ambalaj üzerindeki barkod numarasını girin."}
          </div>
          <input
            value={manual}
            onChange={e=>{ setManual(e.target.value); setManualErr(""); }}
            onKeyDown={e=>e.key==="Enter"&&handleManualSearch()}
            placeholder="Örn: 3282770203745"
            inputMode="numeric"
            autoFocus
            style={{width:"100%",padding:"12px",border:`1.5px solid ${manualErr?"#C0392B":"#E0E0E0"}`,
              borderRadius:10,fontSize:16,outline:"none",boxSizing:"border-box",
              marginBottom:4,letterSpacing:1,textAlign:"center"}}
          />
          {manualErr&&<div style={{color:"#C0392B",fontSize:11,marginBottom:6}}>⚠ {manualErr}</div>}
          <div style={{marginBottom:10}}/>
          <button onClick={handleManualSearch}
            style={{width:"100%",background:"#2C4A3E",color:"#fff",border:"none",borderRadius:10,
              padding:"12px",fontSize:14,fontWeight:700,cursor:"pointer"}}>
            {mode==="fill"?"Kaydet →":"Ara →"}
          </button>
          {status==="manual"&&("BarcodeDetector" in window)&&(
            <button onClick={retryCamera}
              style={{width:"100%",marginTop:8,background:"#F0F0F0",color:"#555",border:"none",
                borderRadius:10,padding:"10px",fontSize:12,cursor:"pointer"}}>
              ← Kameraya dön
            </button>
          )}
        </div>
      )}

      <style>{`@keyframes scan{0%,100%{opacity:.2}50%{opacity:1}}`}</style>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════
// PIN GATE
// ══════════════════════════════════════════════════════════════
function AuthField({label,type="text",value,onChange,placeholder,autoComplete}){
  return(
    <div style={{marginBottom:10,textAlign:"left"}}>
      <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:4}}>{label}</div>
      <input type={type} value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder}
        autoComplete={autoComplete}
        style={{width:"100%",padding:"11px 12px",border:"1.5px solid #E0E0E0",borderRadius:10,
          fontSize:14,outline:"none",boxSizing:"border-box"}}/>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════
// ECZANE GİRİŞ / KAYIT TALEBİ EKRANI
// ══════════════════════════════════════════════════════════════
function PharmacyAuthGate({onLogin,onRegister,onSuccess,onBack}){
  const [mode,setMode]=useState("login"); // login | register
  const [loginForm,setLoginForm]=useState({email:"",password:""});
  const [regForm,setRegForm]=useState({pharmacyName:"",contactName:"",email:"",phone:"",password:"",password2:""});
  const [err,setErr]=useState("");
  const [info,setInfo]=useState("");
  const [busy,setBusy]=useState(false);

  const switchMode=m=>{setMode(m);setErr("");};

  const doLogin=async()=>{
    setErr("");
    if(!loginForm.email.trim()||!loginForm.password){setErr("E-posta ve şifre girin.");return;}
    setBusy(true);
    const res=await onLogin(loginForm.email.trim().toLowerCase(),loginForm.password);
    setBusy(false);
    if(res.ok)onSuccess(res.account);else setErr(res.error);
  };

  const doRegister=async()=>{
    setErr("");
    if(!regForm.pharmacyName.trim()||!regForm.contactName.trim()){setErr("Eczane adı ve yetkili adı gerekli.");return;}
    if(!EMAIL_RE.test(regForm.email.trim())){setErr("Geçerli bir e-posta adresi girin.");return;}
    if(!/^[0-9+()\s-]{7,}$/.test(regForm.phone.trim())){setErr("Geçerli bir telefon numarası girin.");return;}
    if(regForm.password.length<6){setErr("Şifre en az 6 karakter olmalı.");return;}
    if(regForm.password!==regForm.password2){setErr("Şifreler eşleşmiyor.");return;}
    setBusy(true);
    const res=await onRegister({...regForm,email:regForm.email.trim().toLowerCase()});
    setBusy(false);
    if(!res.ok){setErr(res.error);return;}
    setRegForm({pharmacyName:"",contactName:"",email:"",phone:"",password:"",password2:""});
    setInfo("✓ Kayıt talebiniz alındı. E-posta adresinize bir doğrulama bağlantısı gönderilmiş olabilir. Süper admin onayladıktan sonra giriş yapabilirsiniz.");
    setMode("login");
  };

  return(
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",
      background:"#F5F5F5",fontFamily:"'Inter',sans-serif",padding:"24px 0"}}>
      <div style={{background:"#fff",borderRadius:20,padding:32,maxWidth:380,width:"90%",
        boxShadow:"0 4px 24px rgba(0,0,0,.08)"}}>
        <div style={{textAlign:"center",marginBottom:20}}>
          <div style={{fontSize:40,marginBottom:8}}>🏪</div>
          <div style={{fontSize:19,fontWeight:700,color:"#1C1C1A"}}>Eczane Yöneticisi</div>
        </div>
        <div style={{display:"flex",background:"#F0F0F0",borderRadius:10,padding:3,marginBottom:20}}>
          {[{id:"login",l:"Giriş Yap"},{id:"register",l:"Kayıt Talebi"}].map(t=>(
            <button key={t.id} onClick={()=>switchMode(t.id)}
              style={{flex:1,border:"none",borderRadius:8,padding:"9px 0",fontSize:13,cursor:"pointer",
                background:mode===t.id?"#1C1C1A":"transparent",color:mode===t.id?"#fff":"#666",
                fontWeight:mode===t.id?700:500}}>{t.l}</button>
          ))}
        </div>

        {info&&<div style={{background:"#E9F7EF",color:"#1E6B3E",fontSize:12,padding:"10px 12px",
          borderRadius:9,marginBottom:14,lineHeight:1.5}}>{info}</div>}

        {mode==="login"?(
          <>
            <AuthField label="E-posta" type="email" value={loginForm.email}
              onChange={v=>setLoginForm(p=>({...p,email:v}))} placeholder="eczane@ornek.com" autoComplete="username"/>
            <AuthField label="Şifre" type="password" value={loginForm.password}
              onChange={v=>setLoginForm(p=>({...p,password:v}))} placeholder="••••••••" autoComplete="current-password"/>
            {err&&<div style={{color:"#C0392B",fontSize:12,margin:"2px 0 10px"}}>⚠ {err}</div>}
            <button onClick={doLogin} disabled={busy} style={{width:"100%",background:busy?"#999":"#1C1C1A",color:"#fff",border:"none",
              borderRadius:12,padding:14,fontSize:15,fontWeight:700,cursor:busy?"default":"pointer",marginTop:4,
              display:"flex",alignItems:"center",justifyContent:"center",gap:8}}>
              {busy&&<Spinner light size={15}/>}{busy?"Giriş yapılıyor…":"Giriş →"}</button>
          </>
        ):(
          <>
            <AuthField label="Eczane Adı" value={regForm.pharmacyName}
              onChange={v=>setRegForm(p=>({...p,pharmacyName:v}))} placeholder="Merkez Eczanesi"/>
            <AuthField label="Yetkili Ad Soyad" value={regForm.contactName}
              onChange={v=>setRegForm(p=>({...p,contactName:v}))} placeholder="Ad Soyad"/>
            <AuthField label="E-posta" type="email" value={regForm.email}
              onChange={v=>setRegForm(p=>({...p,email:v}))} placeholder="eczane@ornek.com" autoComplete="username"/>
            <AuthField label="Telefon" type="tel" value={regForm.phone}
              onChange={v=>setRegForm(p=>({...p,phone:v}))} placeholder="05xx xxx xx xx"/>
            <AuthField label="Şifre" type="password" value={regForm.password}
              onChange={v=>setRegForm(p=>({...p,password:v}))} placeholder="En az 6 karakter" autoComplete="new-password"/>
            <AuthField label="Şifre (tekrar)" type="password" value={regForm.password2}
              onChange={v=>setRegForm(p=>({...p,password2:v}))} placeholder="Şifrenizi tekrar girin" autoComplete="new-password"/>
            {err&&<div style={{color:"#C0392B",fontSize:12,margin:"2px 0 10px"}}>⚠ {err}</div>}
            <button onClick={doRegister} disabled={busy} style={{width:"100%",background:busy?"#999":"#1C1C1A",color:"#fff",border:"none",
              borderRadius:12,padding:14,fontSize:15,fontWeight:700,cursor:busy?"default":"pointer",marginTop:4,
              display:"flex",alignItems:"center",justifyContent:"center",gap:8}}>
              {busy&&<Spinner light size={15}/>}{busy?"Gönderiliyor…":"Kayıt Talebi Gönder →"}</button>
            <div style={{fontSize:11,color:"#999",marginTop:8,lineHeight:1.5}}>
              Talebiniz süper admin tarafından onaylanana kadar giriş yapamazsınız.
            </div>
          </>
        )}
        <button onClick={onBack} style={{background:"none",border:"none",color:"#999",fontSize:12,
          cursor:"pointer",display:"block",margin:"16px auto 0"}}>← Geri dön</button>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════
// SÜPER ADMİN GİRİŞ EKRANI
// ══════════════════════════════════════════════════════════════
function SuperAdminAuthGate({onLogin,onSuccess,onBack}){
  const [form,setForm]=useState({email:"",password:""});
  const [err,setErr]=useState("");
  const [busy,setBusy]=useState(false);

  const doLogin=async()=>{
    setErr("");
    if(!form.email.trim()||!form.password){setErr("E-posta ve şifre girin.");return;}
    setBusy(true);
    const res=await onLogin(form.email.trim().toLowerCase(),form.password);
    setBusy(false);
    if(res.ok)onSuccess();else setErr(res.error);
  };

  return(
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",
      background:"#F5F5F5",fontFamily:"'Inter',sans-serif"}}>
      <div style={{background:"#fff",borderRadius:20,padding:36,maxWidth:360,width:"90%",
        boxShadow:"0 4px 24px rgba(0,0,0,.08)",textAlign:"center"}}>
        <div style={{fontSize:48,marginBottom:12}}>🔐</div>
        <div style={{fontSize:20,fontWeight:700,color:"#1C1C1A",marginBottom:4}}>Süper Admin</div>
        <div style={{fontSize:13,color:"#888",marginBottom:24}}>E-posta ve şifrenizle giriş yapın</div>
        <AuthField label="E-posta" type="email" value={form.email}
          onChange={v=>setForm(p=>({...p,email:v}))} placeholder="admin@eczane.com" autoComplete="username"/>
        <AuthField label="Şifre" type="password" value={form.password}
          onChange={v=>setForm(p=>({...p,password:v}))} placeholder="••••••••" autoComplete="current-password"/>
        {err&&<div style={{color:"#C0392B",fontSize:12,margin:"2px 0 10px"}}>⚠ {err}</div>}
        <button onClick={doLogin} disabled={busy} style={{width:"100%",background:busy?"#999":"#1C1C1A",color:"#fff",border:"none",
          borderRadius:12,padding:14,fontSize:15,fontWeight:700,cursor:busy?"default":"pointer",marginBottom:8,marginTop:4,
          display:"flex",alignItems:"center",justifyContent:"center",gap:8}}>
          {busy&&<Spinner light size={15}/>}{busy?"Giriş yapılıyor…":"Giriş →"}
        </button>
        <button onClick={onBack} style={{background:"none",border:"none",color:"#999",fontSize:12,cursor:"pointer"}}>
          ← Geri dön
        </button>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════
// HESAP BİLGİLERİ / ŞİFRE DEĞİŞTİRME MODAL
// ══════════════════════════════════════════════════════════════
function ChangeCredentialsModal({title,initialEmail,initialPhone,showPhone=true,onSave,onClose}){
  const [curPass,setCurPass]=useState("");
  const [email,setEmail]=useState(initialEmail||"");
  const [phone,setPhone]=useState(initialPhone||"");
  const [newPass,setNewPass]=useState("");
  const [confPass,setConfPass]=useState("");
  const [err,setErr]=useState("");
  const [busy,setBusy]=useState(false);
  const [notice,setNotice]=useState("");

  const save=async()=>{
    setErr("");
    if(!curPass){setErr("Mevcut şifrenizi girin.");return;}
    if(!EMAIL_RE.test(email.trim())){setErr("Geçerli bir e-posta girin.");return;}
    if(showPhone&&!/^[0-9+()\s-]{7,}$/.test(phone.trim())){setErr("Geçerli bir telefon girin.");return;}
    if(newPass&&newPass.length<6){setErr("Yeni şifre en az 6 karakter olmalı.");return;}
    if(newPass!==confPass){setErr("Yeni şifreler eşleşmiyor.");return;}
    setBusy(true);
    const res=await onSave({currentPassword:curPass,email:email.trim().toLowerCase(),phone:phone.trim(),newPassword:newPass||undefined});
    setBusy(false);
    if(res&&res.ok===false){setErr(res.error);return;}
    if(res&&res.notice){setNotice(res.notice);return;}
    onClose();
  };

  return(
    <div style={{position:"fixed",inset:0,zIndex:500,background:"rgba(0,0,0,.5)",
      display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'Inter',sans-serif"}}>
      <div style={{background:"#fff",borderRadius:20,padding:28,maxWidth:360,width:"90%",maxHeight:"88vh",overflowY:"auto"}}>
        <div style={{fontSize:18,fontWeight:700,color:"#1C1C1A",marginBottom:18}}>👤 {title||"Hesap Bilgileri"}</div>
        <AuthField label="E-posta" type="email" value={email} onChange={setEmail} placeholder="ornek@eczane.com"/>
        {showPhone&&<AuthField label="Telefon" type="tel" value={phone} onChange={setPhone} placeholder="05xx xxx xx xx"/>}
        <div style={{height:1,background:"#F0F0F0",margin:"12px 0"}}/>
        <AuthField label="Mevcut Şifre *" type="password" value={curPass} onChange={setCurPass} placeholder="Değişiklik için gerekli"/>
        <AuthField label="Yeni Şifre (opsiyonel)" type="password" value={newPass} onChange={setNewPass} placeholder="Boş bırakılırsa değişmez"/>
        <AuthField label="Yeni Şifre (tekrar)" type="password" value={confPass} onChange={setConfPass} placeholder="Yeni şifreyi tekrar girin"/>
        {err&&<div style={{color:"#C0392B",fontSize:12,marginBottom:10}}>⚠ {err}</div>}
        {notice&&<div style={{background:"#E9F7EF",color:"#1E6B3E",fontSize:11,padding:"9px 10px",
          borderRadius:8,marginBottom:10,lineHeight:1.5}}>ℹ {notice}</div>}
        <div style={{display:"flex",gap:8,marginTop:8}}>
          <button onClick={onClose}
            style={{flex:1,background:"#F0F0F0",color:"#333",border:"none",borderRadius:10,
              padding:"11px",fontSize:13,cursor:"pointer"}}>{notice?"Kapat":"İptal"}</button>
          <button onClick={notice?onClose:save} disabled={busy}
            style={{flex:1,background:busy?"#999":"#1C1C1A",color:"#fff",border:"none",borderRadius:10,
              padding:"11px",fontSize:13,fontWeight:700,cursor:busy?"default":"pointer",
              display:"flex",alignItems:"center",justifyContent:"center",gap:7}}>
            {busy&&<Spinner light size={13}/>}{busy?"Kaydediliyor…":notice?"Tamam":"Kaydet"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════
// ÜRÜN DETAY MODAL
// ══════════════════════════════════════════════════════════════
function ProductModal({product,qty=0,onAdd,onRemove,onClose,theme}){
  const T=theme||DT;
  const price=product.discountedPrice||product.price||product.basePrice;
  const orig=product.price||product.basePrice;
  const hasDisc=product.discountedPrice&&product.discountedPrice<orig;
  const inBasket=qty>0;
  const stockNum=stockQty(product);
  const atMax=stockNum!=null&&qty>=stockNum;
  const canAdd=hasStock(product)&&!atMax;
  return(
    <div style={{position:"fixed",inset:0,zIndex:200,display:"flex",flexDirection:"column",justifyContent:"flex-end"}}>
      <div onClick={onClose} style={{position:"absolute",inset:0,background:"rgba(0,0,0,.5)"}}/>
      <div style={{position:"relative",background:"#fff",borderRadius:"20px 20px 0 0",
        padding:"20px 16px 40px",maxHeight:"88vh",overflowY:"auto"}}>
        <div style={{width:40,height:4,borderRadius:2,background:"#DDD",margin:"0 auto 16px"}}/>
        {product.photo
          ?<img src={product.photo} alt={product.name}
              style={{width:"100%",height:170,objectFit:"cover",borderRadius:12,marginBottom:14}}
              onError={e=>e.target.style.display="none"}/>
          :<div style={{width:"100%",height:80,background:BRAND_COLORS[product.brand]?.bg||"#F0F0F0",
              borderRadius:12,marginBottom:14,display:"flex",alignItems:"center",justifyContent:"center"}}>
            <Avatar brand={product.brand} size={56}/></div>}
        <div style={{fontSize:11,color:"#888"}}>{product.brand}</div>
        <div style={{fontSize:20,fontWeight:700,color:"#1C1C1A",lineHeight:1.2,marginBottom:8}}>{product.name}</div>
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
          <span style={{fontSize:22,fontWeight:700,color:T.primary}}>{Math.round(price).toLocaleString("tr-TR")} ₺</span>
          {hasDisc&&<>
            <span style={{fontSize:14,color:"#999",textDecoration:"line-through"}}>{Math.round(orig).toLocaleString("tr-TR")} ₺</span>
            <span style={{background:"#FDECEA",color:"#C0392B",fontSize:11,fontWeight:700,padding:"2px 7px",borderRadius:5}}>
              -%{Math.round((1-price/orig)*100)}</span></>}
        </div>
        <div style={{display:"flex",flexWrap:"wrap",gap:4,marginBottom:10}}>
          <BoycottBadge status={product.boycott}/>
          <PregBadge status={product.pregnancy_safe}/>
          {!hasStock(product)&&<span style={{background:"#FFF3E0",color:"#7D4700",fontSize:10,fontWeight:600,padding:"2px 8px",borderRadius:5}}>Stok Yok</span>}
          {(product.certs||[]).map(c=><CertBadge key={c} id={c}/>)}
        </div>
        {product.actives?.length>0&&<div style={{fontSize:12,color:"#555",marginBottom:10}}>
          <b>Aktif içerikler: </b>{product.actives.join(", ")}</div>}
        <div style={{background:"#F8F8F8",borderRadius:10,padding:12,marginBottom:14}}>
          <div style={{fontSize:13,color:"#333",lineHeight:1.6,marginBottom:6}}>{product.desc}</div>
          <div style={{fontSize:12,color:"#666",lineHeight:1.6}}><b>Nasıl kullanılır: </b>{product.how_to}</div>
          <div style={{fontSize:11,color:"#888",marginTop:4}}>
            {product.usage==="am"?"☀️ Sabah":product.usage==="pm"?"🌙 Akşam":"☀️🌙 Sabah & Akşam"}
            {" · "}{CAT_LABELS[product.category]}
          </div>
        </div>
        {product.barcode&&<div style={{marginBottom:14}}><Barcode value={product.barcode} height={30}/></div>}
        {inBasket?(
          <div style={{display:"flex",alignItems:"center",gap:12,justifyContent:"center",
            background:T.bg,border:`1.5px solid ${T.primary}`,borderRadius:12,padding:10}}>
            <button onClick={()=>onRemove(product.id)}
              style={{background:"none",border:"none",color:T.primary,width:40,height:36,
                fontSize:20,fontWeight:700,cursor:"pointer"}}>−</button>
            <span style={{minWidth:24,textAlign:"center",fontSize:17,fontWeight:700,color:"#1C1C1A"}}>{qty}</span>
            <button onClick={()=>canAdd&&onAdd(product)}
              style={{background:"none",border:"none",color:canAdd?T.primary:"#CCC",width:40,height:36,
                fontSize:20,fontWeight:700,cursor:canAdd?"pointer":"default"}}>+</button>
          </div>
        ):(
          <button onClick={()=>{hasStock(product)&&onAdd(product);}}
            style={{width:"100%",background:hasStock(product)?T.primary:"#DDD",
              color:"#fff",border:"none",borderRadius:12,padding:14,
              fontSize:15,fontWeight:700,cursor:hasStock(product)?"pointer":"default"}}>
            {hasStock(product)?"Sepete Ekle →":"Stok Yok"}
          </button>
        )}
        {atMax&&<div style={{textAlign:"center",fontSize:11,color:"#7D4700",marginTop:6}}>Stok limiti: {stockNum} adet</div>}
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════
// ÜRÜN KARTI
// ══════════════════════════════════════════════════════════════
function ProductCard({product,qty=0,onAdd,onRemove,onDetail,compact=false,recReason=null,theme,isTop=false}){
  const T=theme||DT;
  const inBasket=qty>0;
  const stockNum=stockQty(product);
  const atMax=stockNum!=null&&qty>=stockNum;
  const canAdd=hasStock(product)&&!atMax;
  const price=product.discountedPrice||product.price||product.basePrice;
  const orig=product.price||product.basePrice;
  const hasDisc=product.discountedPrice&&product.discountedPrice<orig;
  return(
    <div style={{position:"relative",background:T.surface,borderRadius:14,
      border:`1.5px solid ${inBasket?T.primary:isTop?"#C4974A":"#E2DED6"}`,
      padding:compact?"10px 12px":"14px",opacity:hasStock(product)?1:.65}}>
      {isTop&&<div style={{position:"absolute",top:-1,right:10,background:"#C4974A",
        color:"#fff",fontSize:9,fontWeight:700,padding:"2px 8px",borderRadius:"0 0 6px 6px",letterSpacing:.3}}>
        ⭐ Sizin için
      </div>}
      <div style={{display:"flex",gap:10,alignItems:"flex-start"}}>
        <div onClick={()=>onDetail&&onDetail(product)} style={{cursor:"pointer",flexShrink:0}}>
          {product.photo
            ?<img src={product.photo} alt={product.name}
                style={{width:compact?36:44,height:compact?36:44,borderRadius:10,objectFit:"cover"}}
                onError={e=>e.target.style.display="none"}/>
            :<Avatar brand={product.brand} size={compact?36:44}/>}
        </div>
        <div style={{flex:1,minWidth:0,cursor:"pointer"}} onClick={()=>onDetail&&onDetail(product)}>
          <div style={{fontSize:10,color:"#888"}}>{product.brand}</div>
          <div style={{fontSize:compact?13:14,fontWeight:600,color:"#1C1C1A",lineHeight:1.3}}>{product.name}</div>
          <div style={{display:"flex",alignItems:"center",gap:5,marginTop:3}}>
            <span style={{fontSize:14,fontWeight:700,color:T.primary}}>{Math.round(price).toLocaleString("tr-TR")} ₺</span>
            {hasDisc&&<>
              <span style={{fontSize:10,color:"#999",textDecoration:"line-through"}}>{Math.round(orig).toLocaleString("tr-TR")} ₺</span>
              <span style={{background:"#FDECEA",color:"#C0392B",fontSize:9,fontWeight:700,padding:"1px 5px",borderRadius:4}}>
                -%{Math.round((1-price/orig)*100)}</span></>}
          </div>
        </div>
        {inBasket?(
          <div style={{display:"flex",alignItems:"center",gap:2,flexShrink:0,background:T.bg,
            borderRadius:10,border:`1px solid ${T.primary}`}}>
            <button onClick={()=>onRemove(product.id)}
              style={{background:"none",border:"none",color:T.primary,width:28,height:30,
                fontSize:16,fontWeight:700,cursor:"pointer"}}>−</button>
            <span style={{minWidth:16,textAlign:"center",fontSize:13,fontWeight:700,color:"#1C1C1A"}}>{qty}</span>
            <button onClick={()=>canAdd&&onAdd(product)}
              style={{background:"none",border:"none",color:canAdd?T.primary:"#CCC",width:28,height:30,
                fontSize:16,fontWeight:700,cursor:canAdd?"pointer":"default"}}>+</button>
          </div>
        ):(
          <button onClick={()=>hasStock(product)&&onAdd(product)}
            style={{background:hasStock(product)?T.primary:"#DDD",color:"#fff",
              border:"none",borderRadius:10,padding:"7px 13px",fontSize:12,fontWeight:600,
              cursor:hasStock(product)?"pointer":"default",flexShrink:0}}>
            {hasStock(product)?"+ Ekle":"Yok"}
          </button>
        )}
      </div>
      <div style={{display:"flex",flexWrap:"wrap",gap:4,marginTop:7}}>
        <BoycottBadge status={product.boycott}/>
        <PregBadge status={product.pregnancy_safe}/>
        {!hasStock(product)&&<span style={{background:"#FFF3E0",color:"#7D4700",fontSize:10,fontWeight:600,padding:"2px 7px",borderRadius:5}}>Stok Yok</span>}
        {atMax&&<span style={{background:"#FFF3E0",color:"#7D4700",fontSize:10,fontWeight:600,padding:"2px 7px",borderRadius:5}}>Stok limiti: {stockNum}</span>}
        {(product.certs||[]).slice(0,2).map(c=><CertBadge key={c} id={c}/>)}
      </div>
      {product.actives?.length>0&&<div style={{fontSize:10,color:"#888",marginTop:4}}>
        <b>Aktif: </b>{product.actives.join(", ")}</div>}
      {recReason&&<div style={{marginTop:5,fontSize:11,color:"#C4974A",fontStyle:"italic"}}>📌 {recReason}</div>}
      {!compact&&<button onClick={()=>onDetail&&onDetail(product)}
        style={{background:"none",border:"none",color:T.accent,fontSize:11,fontWeight:600,
          cursor:"pointer",padding:"5px 0 0",display:"block"}}>Detay ve kullanım →</button>}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════
// MÜŞTERİ KİOSK
// ══════════════════════════════════════════════════════════════
function CustomerKiosk({pharmacyCatalog,pharmacySettings,onAdminTrigger,onCompleteSale}){
  const T=useTheme(pharmacySettings);
  const [basket,setBasket]=useState([]);
  const [tab,setTab]=useState("catalog");
  const [search,setSearch]=useState("");
  const [filterCat,setFilterCat]=useState("hepsi");
  const [filterBoycott,setFilterBoycott]=useState(false);
  const [recs,setRecs]=useState([]);
  const [detailProduct,setDetailProduct]=useState(null);
  const [completed,setCompleted]=useState(false);
  const [showScanner,setShowScanner]=useState(false);
  const [scanMsg,setScanMsg]=useState("");
  const [showQuiz,setShowQuiz]=useState(true);
  const [skinProfile,setSkinProfile]=useState(null);
  const [aiDiscount,setAiDiscount]=useState(null);
  const [discountApplied,setDiscountApplied]=useState(false);
  const logoTaps=useRef(0);const logoTimer=useRef(null);

  // 90 saniye etkileşim olmazsa: sepeti/anketi/detayı sıfırlayıp başa dön (bir sonraki müşteri temiz ekran görsün)
  const resetToIdle=useCallback(()=>{
    setBasket([]);setRecs([]);setAiDiscount(null);setDiscountApplied(false);
    setDetailProduct(null);setCompleted(false);setShowQuiz(true);setSkinProfile(null);
    setTab("catalog");setSearch("");setScanMsg("");
  },[]);
  useIdleTimer(90000,resetToIdle,true);

  const handleLogoTap=()=>{
    logoTaps.current++;clearTimeout(logoTimer.current);
    logoTimer.current=setTimeout(()=>{logoTaps.current=0;},2000);
    if(logoTaps.current>=5){logoTaps.current=0;onAdminTrigger();}
  };

  const basketIds=new Set(basket.map(p=>p.id));
  const basketCount=basket.reduce((s,p)=>s+(p.qty||1),0);
  const totalOrig=basket.reduce((s,p)=>s+(p.price||p.basePrice)*(p.qty||1),0);
  const totalDisc=basket.reduce((s,p)=>s+(p.discountedPrice||p.price||p.basePrice)*(p.qty||1),0);

  // Filtrele + hamilelik güvenliği uygula + profil skoru ile sırala
  const isPregnant=skinProfile&&(skinProfile.pregnant==="pregnant"||skinProfile.pregnant==="breastfeeding"||skinProfile.pregnant==="unsure");
  const filteredBase=pharmacyCatalog
    .filter(p=>hasStock(p))
    .filter(p=>filterCat==="hepsi"||p.category===filterCat)
    .filter(p=>!filterBoycott||p.boycott!=="boykot")
    .filter(p=>!isPregnant||p.pregnancy_safe!=="avoid")
    .filter(p=>search===""||
      p.name.toLowerCase().includes(search.toLowerCase())||
      p.brand.toLowerCase().includes(search.toLowerCase())||
      (p.actives||[]).some(a=>a.toLowerCase().includes(search.toLowerCase())));

  const filtered=skinProfile
    ?[...filteredBase].sort((a,b)=>scoreProduct(b,skinProfile)-scoreProduct(a,skinProfile))
    :filteredBase;

  const updateRecs=useCallback((cb)=>{
    setRecs(getAlgoRecs(cb,pharmacyCatalog,skinProfile));
  },[pharmacyCatalog,skinProfile]);

  // Eczacının Ayarlar panelinden kontrol ettiği sepet indirimi kuralları
  const discCfg={
    enabled: pharmacySettings?.cartDiscountEnabled!==false,
    threshold: Math.max(1, parseInt(pharmacySettings?.cartDiscountThreshold)||1),
    pct: Math.max(1, parseInt(pharmacySettings?.cartDiscountPct)||10),
  };

  const applyDiscToBasket=(nb,pct)=>{
    const rate=1-(pct/100);
    return nb.map(x=>({...x,discountedPrice:Math.round((x.price||x.basePrice)*rate)}));
  };

  // Ürün sepette zaten varsa adedini +1 artırır (stok limiti içinde), yoksa 1 adetle ekler
  const addToBasket=useCallback(p=>{
    const existing=basket.find(b=>b.id===p.id);
    const stockNum=stockQty(p);
    if(!hasStock(p))return;
    if(existing){
      if(stockNum!=null&&existing.qty>=stockNum)return; // stok limiti aşılmasın
    }
    const nb=existing
      ?basket.map(b=>b.id===p.id?{...b,qty:b.qty+1}:b)
      :[...basket,{...p,qty:1}];
    updateRecs(nb);

    const totalQty=nb.reduce((s,x)=>s+x.qty,0);
    let shouldApply=discountApplied;
    let pct=aiDiscount?.discount_pct;

    if(!discountApplied&&discCfg.enabled&&totalQty>=discCfg.threshold){
      const avgMargin=nb.reduce((s,x)=>{
        const pr=x.price||x.basePrice;const c=x.cost;
        return s+(c?(pr-c)/pr:0.3);
      },0)/nb.length;
      if(avgMargin>MIN_MARGIN+0.05){shouldApply=true;pct=discCfg.pct;}
    }

    if(shouldApply&&pct){
      setBasket(applyDiscToBasket(nb,pct));
      if(!discountApplied){
        setAiDiscount({discount_pct:pct,reason:`${totalQty} ürün alımında %${pct} indirim otomatik uygulandı`});
        setDiscountApplied(true);
      }
    } else {
      setBasket(nb);
    }
  },[basket,discountApplied,aiDiscount,updateRecs,discCfg.enabled,discCfg.threshold,discCfg.pct]);

  // Ürünün adedini -1 azaltır; adet 0'a inince sepetten tamamen kaldırılır
  const removeFromBasket=id=>{
    const existing=basket.find(b=>b.id===id);
    if(!existing)return;
    const nb=existing.qty>1
      ?basket.map(b=>b.id===id?{...b,qty:b.qty-1}:b)
      :basket.filter(b=>b.id!==id);
    updateRecs(nb);
    const totalQty=nb.reduce((s,x)=>s+x.qty,0);
    if(discountApplied&&totalQty<discCfg.threshold){
      // İndirim koşulu artık sağlanmıyor — indirim geri alınır
      setBasket(nb.map(x=>({...x,discountedPrice:null})));
      setAiDiscount(null);setDiscountApplied(false);
    } else {
      setBasket(nb);
    }
  };
  const clearBasket=()=>{setBasket([]);setRecs([]);setAiDiscount(null);setDiscountApplied(false);};

  const revertDiscount=()=>{
    setBasket(prev=>prev.map(p=>({...p,discountedPrice:null})));
    setAiDiscount(null);setDiscountApplied(false);
  };

  const handleScanFound=p=>{
    const cp=pharmacyCatalog.find(x=>x.id===p.id&&hasStock(x));
    if(cp){addToBasket(cp);setScanMsg(`✓ ${cp.brand} ${cp.name} sepete eklendi`);setTab("catalog");}
    else setScanMsg("❕ Bu ürün eczane kataloğunda yok");
    setTimeout(()=>setScanMsg(""),3000);
  };

  if(completed)return(
    <div className="kiosk-shell" style={{background:T.bg,minHeight:"100vh",
      display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",
      padding:24,fontFamily:"'Inter',sans-serif",textAlign:"center"}}>
      <div style={{fontSize:64,marginBottom:16}}>✅</div>
      <div style={{fontSize:22,fontWeight:700,color:T.primary,marginBottom:8}}>Listeniz hazır!</div>
      <div style={{fontSize:14,color:"#555",lineHeight:1.6,marginBottom:20}}>
        {basketCount} ürün için eczane görevlisine başvurun.<br/>Ödeme kasada yapılır.
      </div>
      <div style={{background:"#fff",border:"1px solid #E0E0E0",borderRadius:14,
        padding:16,width:"100%",marginBottom:20,textAlign:"left"}}>
        {basket.map(p=>(
          <div key={p.id} style={{display:"flex",justifyContent:"space-between",
            padding:"7px 0",borderBottom:"1px solid #F5F5F5",fontSize:13}}>
            <span>{p.brand} {p.name}</span>
            <span style={{fontWeight:700,color:T.primary}}>
              {Math.round(p.discountedPrice||p.price||p.basePrice).toLocaleString("tr-TR")} ₺
            </span>
          </div>
        ))}
        <div style={{display:"flex",justifyContent:"space-between",padding:"10px 0 0",
          fontSize:16,fontWeight:700,color:T.primary}}>
          <span>Toplam</span><span>{Math.round(totalDisc).toLocaleString("tr-TR")} ₺</span>
        </div>
        {totalDisc<totalOrig&&<div style={{textAlign:"right",fontSize:11,color:"#1E6B3E",marginTop:2}}>
          {Math.round(totalOrig-totalDisc).toLocaleString("tr-TR")} ₺ tasarruf 🎉</div>}
      </div>
      <button onClick={()=>{clearBasket();setCompleted(false);setShowQuiz(true);setSkinProfile(null);setTab("catalog");}}
        style={{background:T.primary,color:"#fff",border:"none",borderRadius:12,
          padding:"13px 32px",fontSize:14,fontWeight:700,cursor:"pointer",width:"100%"}}>
        Yeni Arama Başlat
      </button>
    </div>
  );

  return(
    <div className="kiosk-shell" style={{background:T.bg,minHeight:"100vh",
      paddingBottom:80,fontFamily:"'Inter',sans-serif"}}>
      {showQuiz&&<SkinQuiz onComplete={p=>{setSkinProfile(p);setShowQuiz(false);}} theme={T}/>}
      {showScanner&&<BarcodeScanner catalog={INITIAL_POOL} onFound={handleScanFound}
        onClose={()=>setShowScanner(false)} mode="customer"/>}
      {detailProduct&&<ProductModal product={detailProduct} qty={basket.find(p=>p.id===detailProduct.id)?.qty||0}
        onAdd={addToBasket} onRemove={removeFromBasket}
        onClose={()=>setDetailProduct(null)} theme={T}/>}

      {/* HEADER */}
      <div style={{background:T.primary,color:"#fff",padding:"14px 16px 11px",position:"sticky",top:0,zIndex:50}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div onClick={handleLogoTap} style={{cursor:"default",userSelect:"none"}}>
            <div style={{fontSize:17,fontWeight:700}}>{pharmacySettings?.name||"Dermokozmetik Danışmanı"}</div>
            <div style={{fontSize:11,opacity:.5,marginTop:1}}>
              {[...new Set(pharmacyCatalog.map(p=>p.brand))].slice(0,5).join(" · ")}
            </div>
          </div>
          <div style={{display:"flex",gap:8,alignItems:"center"}}>
            <button onClick={()=>setShowScanner(true)}
              style={{background:"rgba(255,255,255,.15)",border:"none",color:"#fff",
                borderRadius:10,padding:"7px 11px",fontSize:16,cursor:"pointer"}}>📷</button>
            <button onClick={()=>setTab(tab==="catalog"?"basket":"catalog")}
              style={{background:basket.length>0?"#C4974A":"rgba(255,255,255,.15)",border:"none",
                borderRadius:11,padding:"7px 13px",color:"#fff",fontSize:13,fontWeight:600,cursor:"pointer"}}>
              🛒{basketCount>0?` ${basketCount}`:""}
            </button>
          </div>
        </div>
        {/* Profil chip */}
        {skinProfile&&(
          <div style={{marginTop:8,display:"flex",alignItems:"center",gap:6}}>
            <div style={{background:"rgba(255,255,255,.15)",borderRadius:20,
              padding:"3px 10px",fontSize:11,display:"flex",alignItems:"center",gap:5,flexWrap:"wrap"}}>
              {skinProfile.skin_type&&<span>
                {skinProfile.skin_type==="oily"?"💧 Yağlı":skinProfile.skin_type==="dry"?"🌵 Kuru":skinProfile.skin_type==="sensitive"?"🌸 Hassas":"✨ Normal"}
              </span>}
              {skinProfile.concern&&<><span style={{opacity:.5}}>·</span>
              <span>{CONCERN_LABELS[skinProfile.concern]||skinProfile.concern}</span></>}
              {isPregnant&&<><span style={{opacity:.5}}>·</span><span>🤰 Hamile filtresi aktif</span></>}
            </div>
            <button onClick={()=>setShowQuiz(true)}
              style={{background:"rgba(255,255,255,.1)",border:"none",color:"rgba(255,255,255,.7)",
                borderRadius:8,padding:"3px 8px",fontSize:10,cursor:"pointer"}}>Değiştir</button>
          </div>
        )}
        {scanMsg&&<div style={{marginTop:6,background:"rgba(255,255,255,.15)",borderRadius:8,
          padding:"6px 10px",fontSize:12,fontWeight:500}}>{scanMsg}</div>}
      </div>

      {tab==="catalog"?(
        <>
          <div style={{padding:"11px 14px 0"}}>
            <input value={search} onChange={e=>setSearch(e.target.value)}
              placeholder="Ürün, marka veya içerik ara…"
              style={{width:"100%",padding:"10px 14px",border:"1.5px solid #E0E0E0",borderRadius:12,
                fontSize:14,background:T.surface,color:"#1C1C1A",outline:"none",boxSizing:"border-box"}}/>
          </div>
          <div style={{padding:"8px 14px",display:"flex",gap:6,overflowX:"auto",scrollbarWidth:"none"}}>
            {["hepsi","temizleyici","serum","nemlendirici","spf","tonik","tedavi"].map(cat=>(
              <button key={cat} onClick={()=>setFilterCat(cat)}
                style={{padding:"5px 12px",borderRadius:20,fontSize:11,fontWeight:500,
                  border:`1.5px solid ${filterCat===cat?T.primary:"#E0E0E0"}`,
                  background:filterCat===cat?T.primary:T.surface,
                  color:filterCat===cat?"#fff":"#888",cursor:"pointer",whiteSpace:"nowrap",flexShrink:0}}>
                {cat==="hepsi"?"Tümü":CAT_LABELS[cat]}
              </button>
            ))}
            <button onClick={()=>setFilterBoycott(!filterBoycott)}
              style={{padding:"5px 12px",borderRadius:20,fontSize:11,fontWeight:500,
                border:`1.5px solid ${filterBoycott?"#1E6B3E":"#E0E0E0"}`,
                background:filterBoycott?"#E9F7EF":T.surface,
                color:filterBoycott?"#1E6B3E":"#888",cursor:"pointer",whiteSpace:"nowrap",flexShrink:0}}>
              ⛔ Boykot Dışı
            </button>
          </div>
          <div style={{padding:"0 14px",display:"flex",flexDirection:"column",gap:8}}>
            <div style={{fontSize:11,color:"#999",marginBottom:2}}>
              {filtered.length} ürün
              {isPregnant&&<span style={{marginLeft:6,color:"#7D4700",fontWeight:600}}>· Hamile güvenli filtre aktif</span>}
              {skinProfile&&!isPregnant&&<span style={{marginLeft:6,color:T.primary,fontWeight:600}}>· Profilinize göre sıralandı</span>}
            </div>
            {filtered.length===0?(
              <div style={{textAlign:"center",padding:"50px 20px",color:"#999"}}>
                <div style={{fontSize:40,marginBottom:12}}>🧴</div>
                {pharmacyCatalog.length===0
                  ?<>Bu eczanenin kataloğu henüz oluşturulmadı.<br/>Lütfen görevliyle iletişime geçin.</>
                  :<>Bu filtrelerle eşleşen ürün bulunamadı.</>}
              </div>
            ):filtered.map((p,i)=>{
              const score=skinProfile?scoreProduct(p,skinProfile):0;
              const isTop=skinProfile&&score>=3&&i<3;
              return <ProductCard key={p.id} product={p} qty={basket.find(b=>b.id===p.id)?.qty||0}
                onAdd={addToBasket} onRemove={removeFromBasket}
                onDetail={setDetailProduct} theme={T} isTop={isTop}/>;
            })}
          </div>
        </>
      ):(
        <div style={{padding:"14px 14px 0"}}>
          {basket.length===0?(
            <div style={{textAlign:"center",padding:"60px 20px",color:"#999"}}>
              <div style={{fontSize:40,marginBottom:12}}>🛒</div>Sepetiniz boş.
              <br/>
              <button onClick={()=>setTab("catalog")}
                style={{marginTop:14,background:T.primary,color:"#fff",border:"none",
                  borderRadius:10,padding:"10px 20px",fontSize:13,fontWeight:600,cursor:"pointer"}}>
                Ürünlere Dön
              </button>
            </div>
          ):(
            <>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                <div style={{fontSize:11,fontWeight:600,color:"#999",textTransform:"uppercase",letterSpacing:.5}}>
                  Seçtiğiniz Ürünler
                </div>
                <button onClick={clearBasket}
                  style={{background:"#FDECEA",color:"#C0392B",border:"none",borderRadius:7,
                    padding:"4px 10px",fontSize:11,fontWeight:600,cursor:"pointer"}}>Tümünü Kaldır</button>
              </div>
              <div style={{display:"flex",flexDirection:"column",gap:7,marginBottom:14}}>
                {basket.map(p=><ProductCard key={p.id} product={p} qty={p.qty||1}
                  onAdd={addToBasket} onRemove={removeFromBasket}
                  onDetail={setDetailProduct} compact theme={T}/>)}
              </div>
              <div style={{background:T.primary,color:"#fff",borderRadius:14,padding:"14px 16px",
                display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
                <div>
                  <div style={{fontSize:11,opacity:.6}}>Toplam ({basketCount} ürün)</div>
                  <div style={{fontSize:24,fontWeight:700}}>{Math.round(totalDisc).toLocaleString("tr-TR")} ₺</div>
                  {totalDisc<totalOrig&&<div style={{fontSize:11,opacity:.7,textDecoration:"line-through"}}>
                    {Math.round(totalOrig).toLocaleString("tr-TR")} ₺</div>}
                </div>
                <div style={{fontSize:11,opacity:.7,textAlign:"right",lineHeight:1.4}}>Ödeme eczanede yapılır</div>
              </div>
              {/* İndirim önerisi — otomatik uygulanır */}
              {aiDiscount&&(
                <div style={{background:"#FFFDE7",border:"1px solid #FBC02D",borderRadius:12,padding:14,marginBottom:14}}>
                  <div style={{fontSize:13,fontWeight:700,color:"#7D4700",marginBottom:4}}>
                    🎁 İndirim Önerisi: %{aiDiscount.discount_pct}
                  </div>
                  <div style={{fontSize:11,color:"#7D4700",lineHeight:1.5,marginBottom:10}}>{aiDiscount.reason}</div>
                  <div style={{display:"flex",gap:8}}>
                    <div style={{flex:1,background:"#E9F7EF",borderRadius:8,padding:"9px",
                      textAlign:"center",fontSize:12,fontWeight:600,color:"#1E6B3E"}}>✓ Uygulandı</div>
                    <button onClick={revertDiscount}
                      style={{background:"#FDECEA",color:"#C0392B",border:"none",borderRadius:8,
                        padding:"9px 12px",fontSize:12,fontWeight:600,cursor:"pointer"}}>Geri Al</button>
                  </div>
                </div>
              )}
              <button onClick={()=>{
                  if(onCompleteSale&&basket.length>0)onCompleteSale({
                    date:new Date().toISOString(),
                    items:basket.map(p=>({id:p.id,name:p.name,brand:p.brand,category:p.category,
                      qty:p.qty||1,
                      unitPrice:p.discountedPrice||p.price||p.basePrice,
                      origPrice:p.price||p.basePrice, cost:p.cost||null})),
                    total:totalDisc, totalOrig, itemCount:basketCount,
                    discountApplied,
                  });
                  setCompleted(true);
                }}
                style={{width:"100%",background:T.primary,color:"#fff",border:"none",borderRadius:12,
                  padding:"14px",fontSize:15,fontWeight:700,cursor:"pointer",marginBottom:16}}>
                Listeyi Tamamla ✓
              </button>
              {recs.length>0&&(
                <>
                  <div style={{fontSize:11,fontWeight:600,color:"#999",marginBottom:8,
                    textTransform:"uppercase",letterSpacing:.5}}>📌 Algoritma Önerisi</div>
                  <div style={{display:"flex",flexDirection:"column",gap:7}}>
                    {recs.map(({prod,reason})=>prod&&(
                      <ProductCard key={prod.id} product={prod} qty={basket.find(b=>b.id===prod.id)?.qty||0}
                        onAdd={addToBasket} onRemove={removeFromBasket}
                        onDetail={setDetailProduct} compact recReason={reason} theme={T}/>
                    ))}
                  </div>
                </>
              )}
            </>
          )}
        </div>
      )}

      {/* ALT NAV */}
      <div className="kiosk-shell" style={{position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",width:"100%",
        background:T.surface,borderTop:"1px solid #E0E0E0",display:"flex"}}>
        {[{id:"catalog",l:"Ürünler",i:"🧴"},{id:"basket",l:basketCount>0?`Sepet (${basketCount})`:"Sepet",i:"🛒"}]
          .map(t=>(
          <button key={t.id} onClick={()=>setTab(t.id)}
            style={{flex:1,border:"none",background:"none",padding:"11px 0 9px",cursor:"pointer",
              color:tab===t.id?T.primary:"#999",fontWeight:tab===t.id?700:400,fontSize:11,
              borderTop:`2.5px solid ${tab===t.id?T.primary:"transparent"}`,
              display:"flex",flexDirection:"column",alignItems:"center",gap:2}}>
            <span style={{fontSize:17}}>{t.i}</span>{t.l}
          </button>
        ))}
        <button onClick={onAdminTrigger}
          style={{width:44,border:"none",background:"none",padding:"11px 0 9px",cursor:"pointer",
            color:"rgba(0,0,0,.08)",fontSize:11,display:"flex",flexDirection:"column",alignItems:"center",gap:2,
            borderTop:"2.5px solid transparent"}}>
          <span style={{fontSize:14}}>⚙</span>
        </button>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════
// EXCEL / CSV İÇE AKTARMA
// Barkod · Ürün Adı · Fiyat sütunlarını otomatik tanır
// ══════════════════════════════════════════════════════════════
function ExcelImport({ masterPool, pharmacyCatalog, setPool, onImport }) {
  const [rows,    setRows]    = useState([]);   // ham satırlar
  const [headers, setHeaders] = useState([]);   // sütun başlıkları
  const [map,     setMap]     = useState({ barcode:"", name:"", price:"", brand:"", stock:"" });
  const [preview, setPreview] = useState([]);   // eşleştirilmiş önizleme
  const [step,    setStep]    = useState("upload"); // upload|map|preview|done
  const [result,  setResult]  = useState({ added:0, updated:0, skipped:0, autoCreated:0, unknown:0, unknownList:[] });
  const [loading, setLoading] = useState(false);
  const [importing, setImporting] = useState(false);
  const [autoCreateUnknown, setAutoCreateUnknown] = useState(true);
  const fileRef = useRef(null);

  // Sütun adı → alan tahmini
  const guessField = (col) => {
    const c = col.toLowerCase().trim();
    if (/barkod|barcode|ean|upc|kod|code/.test(c))         return "barcode";
    if (/ürün|urun|ad|name|product|isim|açıklama|aciklama/.test(c)) return "name";
    if (/fiyat|price|tutar|ücret|ucret|satis/.test(c))     return "price";
    if (/marka|brand/.test(c))                              return "brand";
    if (/stok|stock|adet|miktar|qty|quantity/.test(c))       return "stock";
    return "";
  };

  const handleFile = async (file) => {
    if (!file) return;
    setLoading(true);
    try {
      const buf  = await file.arrayBuffer();
      const wb   = XLSX.read(buf, { type:"array" });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header:1, defval:"" });

      if (!data.length) { alert("Dosya boş."); setLoading(false); return; }

      // İlk satır başlık mı değer mi?
      const firstRow = data[0].map(String);
      const isHeader = firstRow.some(h => /[a-zçğışöü]/i.test(h));
      const hdrs = isHeader ? firstRow : firstRow.map((_,i) => `Sütun ${i+1}`);
      const dataRows = isHeader ? data.slice(1) : data;

      setHeaders(hdrs);
      setRows(dataRows.filter(r => r.some(c => String(c).trim())));

      // Otomatik eşleştirme
      const auto = { barcode:"", name:"", price:"", brand:"", stock:"" };
      hdrs.forEach((h,i) => {
        const f = guessField(h);
        if (f && !auto[f]) auto[f] = String(i);
      });
      setMap(auto);
      setStep("map");
    } catch(e) {
      alert("Dosya okunamadı. Lütfen .xlsx veya .csv formatında yükleyin.\n" + e.message);
    } finally {
      setLoading(false);
    }
  };

  // Önizleme oluştur
  const buildPreview = () => {
    const prev = rows.slice(0, 8).map(row => ({
      barcode: String(row[parseInt(map.barcode)] ?? "").trim(),
      name:    String(row[parseInt(map.name)]    ?? "").trim(),
      price:   parseFloat(String(row[parseInt(map.price)] ?? "").replace(",",".")) || 0,
      brand:   map.brand !== "" ? String(row[parseInt(map.brand)] ?? "").trim() : "",
      stock:   map.stock !== "" ? (parseInt(String(row[parseInt(map.stock)] ?? "").replace(/[^\d-]/g,""),10)) : null,
    })).filter(r => r.name || r.barcode);
    setPreview(prev);
    setStep("preview");
  };

  // İçe aktar
  const doImport = async () => {
    setImporting(true);
    const existingBarcodes = new Map(pharmacyCatalog.map(p => [p.barcode, p]));
    const poolByBarcode = new Map(masterPool.map(p => [p.barcode, p]));
    const createdThisRun = new Map(); // aynı çalıştırmada aynı barkod 2 kere oluşturulmasın

    let added = 0, updated = 0, skipped = 0, autoCreated = 0;
    const newItems = [];
    const unknownList = []; // havuzda bulunamayan, tanımsız ürünler (oto oluşturma kapalıyken)
    const newPoolRows = []; // yeni oluşturulan havuz ürünleri (yerel state'e eklemek için)

    for (const row of rows) {
      const barcode = String(row[parseInt(map.barcode)] ?? "").trim();
      const name    = String(row[parseInt(map.name)]    ?? "").trim();
      const price   = parseFloat(String(row[parseInt(map.price)] ?? "").replace(",",".")) || 0;
      const brand   = map.brand !== "" ? String(row[parseInt(map.brand)] ?? "").trim() : "";
      const stockRaw = map.stock !== "" ? String(row[parseInt(map.stock)] ?? "").replace(/[^\d-]/g,"") : "";
      const stockVal = stockRaw !== "" && !isNaN(parseInt(stockRaw,10)) ? Math.max(0,parseInt(stockRaw,10)) : null;

      if (!name && !barcode) { skipped++; continue; }

      const poolProduct = barcode ? (poolByBarcode.get(barcode) || createdThisRun.get(barcode)) : null;
      const existingCat = barcode ? existingBarcodes.get(barcode) : null;

      if (existingCat) {
        if (price > 0 || stockVal !== null) {
          newItems.push({ ...existingCat,
            price: price > 0 ? price : existingCat.price,
            stock: stockVal !== null ? stockVal : existingCat.stock,
          });
          updated++;
        } else {
          newItems.push(existingCat);
          skipped++;
        }
      } else if (poolProduct) {
        newItems.push({
          ...poolProduct,
          price: price > 0 ? price : poolProduct.basePrice,
          stock: stockVal !== null ? stockVal : true, cost: null, discountedPrice: null,
        });
        added++;
      } else if (autoCreateUnknown && (name || barcode)) {
        // Havuzda yok — hızlı kurulum için otomatik olarak havuza minimal bilgiyle eklenir
        const row_ = {
          barcode, brand: brand || "Bilinmiyor", name: name || "İsimsiz Ürün",
          category: "nemlendirici", usage: "both", base_price: price,
          actives: [], pairs_with: [], certs: [], skin_types: [], concerns: [],
          boycott: "bilinmiyor", pregnancy_safe: "unknown",
          description: "Excel ile hızlı kurulumda otomatik oluşturuldu — incelenmesi önerilir.", how_to: "",
        };
        const { data, error } = await supabase.from("pool").insert(row_).select().single();
        if (!error && data) {
          const created = poolRowToJs(data);
          createdThisRun.set(barcode, created);
          newPoolRows.push(created);
          newItems.push({ ...created, price: price>0?price:created.basePrice,
            stock: stockVal !== null ? stockVal : true, cost: null, discountedPrice: null });
          added++; autoCreated++;
        } else {
          unknownList.push({ barcode, name: name || "(isim yok)", brand });
        }
      } else {
        // Barkod havuzda tanımlı değil — otomatik ürün oluşturulmadı, admin onayına düşer
        unknownList.push({ barcode, name: name || "(isim yok)", brand });
      }
    }

    if (newPoolRows.length > 0 && setPool) {
      setPool(prev => [...prev, ...newPoolRows]);
    }

    // Katalogdaki ürünlerden bu içe aktarmada olmayanları koru
    const importBarcodes = new Set(newItems.map(p => p.barcode).filter(Boolean));
    const kept = pharmacyCatalog.filter(p =>
      !p.barcode || !importBarcodes.has(p.barcode)
    );

    onImport([...kept, ...newItems]);
    setResult({ added, updated, skipped, autoCreated, unknown: unknownList.length, unknownList });
    setImporting(false);
    setStep("done");
  };

  const reset = () => {
    setRows([]); setHeaders([]); setPreview([]);
    setMap({ barcode:"", name:"", price:"", brand:"", stock:"" });
    setStep("upload"); setResult({ added:0, updated:0, skipped:0, autoCreated:0, unknown:0, unknownList:[] });
    if (fileRef.current) fileRef.current.value = "";
  };

  const FIELD_LABELS = { barcode:"Barkod", name:"Ürün Adı", price:"Fiyat", brand:"Marka (opsiyonel)", stock:"Stok Adedi (opsiyonel)" };

  return (
    <div style={{padding:"0 0 20px"}}>
      {/* Upload */}
      {step === "upload" && (
        <div>
          <div style={{background:"#F8F9FA",borderRadius:12,border:"2px dashed #CCC",
            padding:24,textAlign:"center",marginBottom:12}}>
            <div style={{fontSize:32,marginBottom:8}}>📊</div>
            <div style={{fontSize:14,fontWeight:600,color:"#333",marginBottom:4}}>Excel veya CSV yükle</div>
            <div style={{fontSize:12,color:"#888",marginBottom:14,lineHeight:1.5}}>
              Barkod, ürün adı, fiyat ve stok sütunları otomatik tanınır.<br/>
              .xlsx · .xls · .csv formatları desteklenir.
            </div>
            <input
              ref={fileRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={e=>handleFile(e.target.files[0])}
              style={{display:"none"}}
            />
            <button type="button" onClick={()=>fileRef.current&&fileRef.current.click()}
              style={{display:"inline-block",background:"#1C1C1A",color:"#fff",border:"none",
                borderRadius:10,padding:"10px 24px",fontSize:13,fontWeight:600,cursor:"pointer"}}>
              {loading ? "⏳ Yükleniyor…" : "Dosya Seç"}
            </button>
          </div>
          <div style={{background:autoCreateUnknown?"#EEF0FB":"#F4F4F4",borderRadius:10,padding:12,marginBottom:12,
            display:"flex",gap:10,alignItems:"flex-start"}}>
            <input type="checkbox" checked={autoCreateUnknown} onChange={e=>setAutoCreateUnknown(e.target.checked)}
              style={{width:18,height:18,marginTop:2,flexShrink:0,cursor:"pointer",accentColor:"#1C1C1A"}}/>
            <div>
              <div style={{fontSize:12,fontWeight:700,color:autoCreateUnknown?"#3949AB":"#555"}}>
                ⚡ Hızlı kurulum: havuzda olmayan ürünleri otomatik oluştur
              </div>
              <div style={{fontSize:11,color:"#777",marginTop:3,lineHeight:1.5}}>
                Açıkken, listende olup havuzda tanımlı olmayan tüm ürünler minimum bilgiyle otomatik
                oluşturulup direkt kataloğuna eklenir — ilk kurulumda stoğunu tek seferde aktarabilirsin.
                Kapalıyken tanımsız ürünler eklenmez, sadece uyarı olarak listelenir.
              </div>
            </div>
          </div>
          {/* Beklenen format */}
          <div style={{background:"#EEF0FB",borderRadius:10,padding:12}}>
            <div style={{fontSize:11,fontWeight:600,color:"#3949AB",marginBottom:6}}>
              📋 Beklenen sütun formatı (örnek):
            </div>
            <div style={{fontSize:11,fontFamily:"monospace",color:"#555",lineHeight:1.8}}>
              | Barkod &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;| Ürün Adı &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;| Fiyat |<br/>
              | 3282770203745 | Avène Temizleyici | 520 &nbsp;&nbsp;|<br/>
              | 3522930002285 | Caudalie Serum &nbsp;&nbsp;&nbsp;| 1250 &nbsp;|
            </div>
          </div>
        </div>
      )}

      {/* Sütun eşleştirme */}
      {step === "map" && (
        <div>
          <div style={{fontSize:13,fontWeight:600,color:"#1C1C1A",marginBottom:4}}>
            Sütun Eşleştirme
          </div>
          <div style={{fontSize:11,color:"#888",marginBottom:14}}>
            {rows.length} satır okundu. Hangi sütun hangi alana karşılık geliyor?
          </div>

          {Object.entries(FIELD_LABELS).map(([field, label]) => (
            <div key={field} style={{marginBottom:10}}>
              <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:4}}>
                {label}{field!=="brand"&&field!=="stock"&&<span style={{color:"#C0392B"}}> *</span>}
              </div>
              <select
                value={map[field]}
                onChange={e=>setMap(p=>({...p,[field]:e.target.value}))}
                style={{width:"100%",padding:"8px 10px",border:"1px solid #E0E0E0",
                  borderRadius:8,fontSize:13,color:"#333",background:"#fff"}}>
                <option value="">— Seç —</option>
                {headers.map((h,i)=>(
                  <option key={i} value={String(i)}>{h} (örn: {String(rows[0]?.[i]??"")})</option>
                ))}
              </select>
            </div>
          ))}

          <div style={{display:"flex",gap:8,marginTop:16}}>
            <button onClick={reset}
              style={{flex:1,background:"#F0F0F0",color:"#333",border:"none",
                borderRadius:10,padding:"10px",fontSize:13,cursor:"pointer"}}>
              ← Geri
            </button>
            <button
              onClick={buildPreview}
              disabled={!map.name && !map.barcode}
              style={{flex:2,background: map.name||map.barcode ? "#1C1C1A":"#CCC",color:"#fff",
                border:"none",borderRadius:10,padding:"10px",fontSize:13,fontWeight:600,
                cursor:map.name||map.barcode?"pointer":"default"}}>
              Önizle →
            </button>
          </div>
        </div>
      )}

      {/* Önizleme */}
      {step === "preview" && (
        <div>
          <div style={{fontSize:13,fontWeight:600,color:"#1C1C1A",marginBottom:4}}>
            Önizleme — İlk 8 Satır
          </div>
          <div style={{fontSize:11,color:"#888",marginBottom:10}}>
            Toplam {rows.length} satır içe aktarılacak.
          </div>

          <div style={{overflowX:"auto",marginBottom:14}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
              <thead>
                <tr style={{background:"#F5F5F5"}}>
                  {["Barkod","Ürün Adı","Fiyat (₺)","Stok","Durum"].map(h=>(
                    <th key={h} style={{padding:"6px 8px",textAlign:"left",
                      fontWeight:600,color:"#555",whiteSpace:"nowrap",
                      borderBottom:"1px solid #E0E0E0"}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {preview.map((r,i)=>{
                  const inPool = masterPool.find(p=>p.barcode===r.barcode);
                  const inCat  = pharmacyCatalog.find(p=>p.barcode===r.barcode);
                  const status = inCat ? "🔄 Güncelle" : inPool ? "✅ Havuzdan" : "➕ Yeni";
                  const stColor = inCat?"#7D4700":inPool?"#1E6B3E":"#1A237E";
                  return (
                    <tr key={i} style={{borderBottom:"1px solid #F0F0F0"}}>
                      <td style={{padding:"5px 8px",color:"#888",fontFamily:"monospace",fontSize:10}}>
                        {r.barcode||"—"}
                      </td>
                      <td style={{padding:"5px 8px",color:"#1C1C1A",maxWidth:150,overflow:"hidden",
                        textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.name||"—"}</td>
                      <td style={{padding:"5px 8px",fontWeight:600,color:"#2C4A3E"}}>
                        {r.price>0?r.price.toLocaleString("tr-TR")+" ₺":"—"}
                      </td>
                      <td style={{padding:"5px 8px",color:"#555"}}>
                        {r.stock!=null?r.stock+" ad.":"—"}
                      </td>
                      <td style={{padding:"5px 8px",color:stColor,fontWeight:600,whiteSpace:"nowrap"}}>
                        {status}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          <div style={{background:"#EEF0FB",borderRadius:8,padding:10,marginBottom:14,fontSize:11,color:"#333",lineHeight:1.6}}>
            ✅ <b>Havuzdan</b>: Mevcut ürün verileri kullanılır, fiyat güncellenir.<br/>
            🔄 <b>Güncelle</b>: Katalogdaki ürünün fiyatı güncellenir.<br/>
            ➕ <b>Yeni</b>: Minimum bilgiyle kataloga eklenir, sonra düzenleyebilirsiniz.
          </div>

          <div style={{display:"flex",gap:8}}>
            <button onClick={()=>setStep("map")} disabled={importing}
              style={{flex:1,background:"#F0F0F0",color:"#333",border:"none",
                borderRadius:10,padding:"10px",fontSize:13,cursor:importing?"default":"pointer"}}>
              ← Geri
            </button>
            <button onClick={doImport} disabled={importing}
              style={{flex:2,background:importing?"#999":"#2C4A3E",color:"#fff",border:"none",
                borderRadius:10,padding:"10px",fontSize:13,fontWeight:700,cursor:importing?"default":"pointer"}}>
              {importing?"Aktarılıyor…":`${rows.length} Ürünü İçe Aktar →`}
            </button>
          </div>
          {importing&&<div style={{fontSize:10,color:"#999",textAlign:"center",marginTop:6}}>
            Havuza yeni ürün ekleniyorsa biraz sürebilir, lütfen sayfadan ayrılmayın.
          </div>}
        </div>
      )}

      {/* Tamamlandı */}
      {step === "done" && (
        <div style={{textAlign:"center"}}>
          <div style={{fontSize:40,marginBottom:12}}>{result.unknown>0?"⚠️":"🎉"}</div>
          <div style={{fontSize:16,fontWeight:700,color:"#1C1C1A",marginBottom:16}}>
            İçe aktarma tamamlandı
          </div>
          <div style={{display:"flex",gap:8,marginBottom:result.autoCreated>0?8:(result.unknown>0?12:20),flexWrap:"wrap"}}>
            {[
              {label:"Eklendi",value:result.added,bg:"#E9F7EF",c:"#1E6B3E"},
              {label:"Güncellendi",value:result.updated,bg:"#FEF3E0",c:"#7D4700"},
              {label:"Atlandı",value:result.skipped,bg:"#F4F4F4",c:"#666"},
              {label:"Tanımsız",value:result.unknown,bg:result.unknown>0?"#FDECEA":"#F4F4F4",c:result.unknown>0?"#8B2E2E":"#999"},
            ].map(({label,value,bg,c})=>(
              <div key={label} style={{flex:1,minWidth:70,background:bg,borderRadius:10,padding:"12px 6px",textAlign:"center"}}>
                <div style={{fontSize:22,fontWeight:700,color:c}}>{value}</div>
                <div style={{fontSize:10,color:c,marginTop:2}}>{label}</div>
              </div>
            ))}
          </div>
          {result.autoCreated>0&&(
            <div style={{background:"#EEF0FB",color:"#3949AB",borderRadius:10,padding:"9px 12px",
              fontSize:11,marginBottom:result.unknown>0?12:20,textAlign:"left",lineHeight:1.5}}>
              ⚡ Bunlardan <b>{result.autoCreated}</b> tanesi havuzda tanımlı olmadığı için otomatik oluşturuldu.
              Doğruluğu için Havuz'dan (süper admin) gözden geçirilmesi önerilir.
            </div>
          )}
          {result.unknown>0&&(
            <div style={{background:"#FDECEA",borderRadius:10,padding:12,marginBottom:20,textAlign:"left"}}>
              <div style={{fontSize:12,fontWeight:700,color:"#8B2E2E",marginBottom:6}}>
                ⚠ {result.unknown} ürün havuzda tanımlı değil, kataloğunuza eklenmedi
              </div>
              <div style={{fontSize:11,color:"#8B2E2E",lineHeight:1.6,marginBottom:8}}>
                Bu barkodlar master havuzda bulunamadı. Sağlıklı veri için ürünler önce
                süper admin tarafından Havuz'a eklenmeli, sonra kataloğunuza dahil edilebilir.
              </div>
              <div style={{maxHeight:140,overflowY:"auto",display:"flex",flexDirection:"column",gap:4}}>
                {result.unknownList.slice(0,20).map((u,i)=>(
                  <div key={i} style={{background:"#fff",borderRadius:6,padding:"5px 8px",fontSize:10,color:"#555"}}>
                    <span style={{fontFamily:"monospace",color:"#999"}}>{u.barcode||"(barkod yok)"}</span>
                    {" — "}{u.brand&&`${u.brand} `}{u.name}
                  </div>
                ))}
                {result.unknownList.length>20&&
                  <div style={{fontSize:10,color:"#8B2E2E",textAlign:"center",padding:4}}>
                    +{result.unknownList.length-20} tane daha…
                  </div>}
              </div>
            </div>
          )}
          <button onClick={reset}
            style={{width:"100%",background:"#1C1C1A",color:"#fff",border:"none",
              borderRadius:10,padding:"11px",fontSize:13,fontWeight:600,cursor:"pointer"}}>
            Yeni Dosya Yükle
          </button>
        </div>
      )}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════
// ECZANE ADMİN
// ══════════════════════════════════════════════════════════════
function CatalogItem({item,onUpdate,onRemove,selected=false,onToggleSelect,sales=[]}){
  const [price,setPrice]=useState(String(item.price||item.basePrice));
  const [cost,setCost]=useState(String(item.cost||""));
  const [saved,setSaved]=useState(false);

  const save=()=>{
    const p=parseInt(price);const c=cost?parseInt(cost):null;
    if(!isNaN(p)&&p>0)onUpdate(item.id,"price",p);
    onUpdate(item.id,"cost",c);
    setSaved(true);setTimeout(()=>setSaved(false),2000);
  };

  const discPct=item.discountedPrice&&(item.price||item.basePrice)>0
    ?Math.round((1-item.discountedPrice/(item.price||item.basePrice))*100):null;

  const toggleSkinType=(st)=>{
    const cur=item.skin_types||[];
    onUpdate(item.id,"skin_types",cur.includes(st)?cur.filter(x=>x!==st):[...cur,st]);
  };
  const toggleConcern=(c)=>{
    const cur=item.concerns||[];
    onUpdate(item.id,"concerns",cur.includes(c)?cur.filter(x=>x!==c):[...cur,c]);
  };

  const qty=stockQty(item);
  const movement=getMovement(item,sales);
  const daysLeft=daysUntil(item.skt);
  const discSuggest=suggestDiscount(item,sales);
  const URGENCY_META={high:{bg:"#FDECEA",color:"#8B2E2E",label:"🔴 Acil"},medium:{bg:"#FFF3E0",color:"#7D4700",label:"🟠 Öneri"},low:{bg:"#FEFCE8",color:"#7D6608",label:"🟡 Düşük"}};

  const applyAutoDiscount=()=>{
    if(!discSuggest)return;
    const base=item.price||item.basePrice;
    onUpdate(item.id,"discountedPrice",Math.round(base*(1-discSuggest.percent/100)));
  };

  return(
    <div style={{background:"#fff",borderRadius:14,border:`1px solid ${selected?"#1C1C1A":"#E0E0E0"}`,
      padding:14,marginBottom:10,boxShadow:selected?"0 0 0 2px rgba(28,28,26,.08)":"none"}}>
      <div style={{display:"flex",gap:10,alignItems:"center",marginBottom:10}}>
        {onToggleSelect&&<input type="checkbox" checked={selected} onChange={onToggleSelect}
          style={{width:17,height:17,flexShrink:0,cursor:"pointer",accentColor:"#1C1C1A"}}/>}
        {item.photo?<img src={item.photo} style={{width:40,height:40,borderRadius:8,objectFit:"cover",flexShrink:0}}
          onError={e=>e.target.style.display="none"}/>:<Avatar brand={item.brand} size={40}/>}
        <div style={{flex:1}}>
          <div style={{fontSize:10,color:"#999"}}>{item.brand}</div>
          <div style={{fontSize:13,fontWeight:600,color:"#1C1C1A"}}>{item.name}</div>
          <div style={{fontSize:11,color:"#2C4A3E",fontWeight:700}}>
            {(item.discountedPrice||item.price||item.basePrice).toLocaleString("tr-TR")} ₺
            {discPct&&<span style={{marginLeft:6,background:"#FDECEA",color:"#C0392B",fontSize:9,fontWeight:700,padding:"1px 5px",borderRadius:4}}>-%{discPct}</span>}
          </div>
          <div style={{display:"flex",flexWrap:"wrap",gap:4,marginTop:4}}>
            <span style={{background:movement.bg,color:movement.color,fontSize:9,fontWeight:700,padding:"1px 6px",borderRadius:4}}>
              📈 {movement.label}
            </span>
            {qty!=null&&<span style={{background:qty<=0?"#FDECEA":qty<10?"#FFF3E0":"#F4F4F4",
              color:qty<=0?"#8B2E2E":qty<10?"#7D4700":"#666",fontSize:9,fontWeight:700,padding:"1px 6px",borderRadius:4}}>
              📦 {qty} adet</span>}
            {daysLeft!=null&&<span style={{background:daysLeft<0?"#FDECEA":daysLeft<=30?"#FFF3E0":"#F4F4F4",
              color:daysLeft<0?"#8B2E2E":daysLeft<=30?"#7D4700":"#666",fontSize:9,fontWeight:700,padding:"1px 6px",borderRadius:4}}>
              ⏳ {daysLeft<0?"SKT geçti":`${daysLeft} gün SKT`}</span>}
          </div>
        </div>
        <button onClick={()=>onRemove(item.id)}
          style={{background:"#FDECEA",color:"#C0392B",border:"none",borderRadius:8,
            padding:"5px 10px",fontSize:11,fontWeight:600,cursor:"pointer"}}>Çıkar</button>
      </div>

      {item.barcode&&<div style={{background:"#FAFAFA",borderRadius:8,padding:"8px 4px",marginBottom:10}}>
        <Barcode value={item.barcode} height={26}/>
      </div>}

      {/* Oto İndirim Önerisi */}
      {discSuggest&&(()=>{const um=URGENCY_META[discSuggest.urgency];return(
        <div style={{background:um.bg,borderRadius:10,padding:10,marginBottom:10}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
            <span style={{fontSize:11,fontWeight:700,color:um.color}}>{um.label} · 💡 Oto İndirim Önerisi: %{discSuggest.percent}</span>
            <button onClick={applyAutoDiscount}
              style={{background:um.color,color:"#fff",border:"none",borderRadius:7,
                padding:"4px 10px",fontSize:10,fontWeight:700,cursor:"pointer",flexShrink:0}}>Uygula</button>
          </div>
          <div style={{fontSize:10,color:um.color,lineHeight:1.5}}>
            {discSuggest.reasons.map((r,i)=><div key={i}>• {r}</div>)}
          </div>
        </div>
      );})()}

      {/* Stok Adedi */}
      <Row label="Stok Adedi">
        <input type="number" min="0" defaultValue={qty!=null?qty:""} placeholder="—"
          onBlur={e=>{const v=e.target.value===""?true:Math.max(0,parseInt(e.target.value)||0);onUpdate(item.id,"stock",v);}}
          style={{width:80,padding:"4px 8px",border:"1px solid #E0E0E0",borderRadius:7,fontSize:11,
            textAlign:"right",outline:"none"}}/>
      </Row>

      {/* Haftalık Satış Hızı — mümkünse gerçek satış verisinden otomatik */}
      {(() => {
        const computed=computeSalesVelocity(item.id,sales);
        if(computed!=null){
          return(
            <Row label="Haftalık Satış (adet)">
              <span style={{fontSize:11,fontWeight:700,color:"#1A5276",background:"#D6EAF8",
                padding:"3px 8px",borderRadius:6}}>📊 {computed} / hafta · otomatik</span>
            </Row>
          );
        }
        return(
          <Row label="Haftalık Satış (adet, tahmini)">
            <input type="number" min="0" defaultValue={item.weeklySales??""} placeholder="—"
              onBlur={e=>{const v=e.target.value===""?null:Math.max(0,parseInt(e.target.value)||0);onUpdate(item.id,"weeklySales",v);}}
              style={{width:80,padding:"4px 8px",border:"1px solid #E0E0E0",borderRadius:7,fontSize:11,
                textAlign:"right",outline:"none"}}/>
          </Row>
        );
      })()}

      {/* Boykot */}
      <Row label="Boykot">
        <select value={item.boycott} onChange={e=>onUpdate(item.id,"boycott",e.target.value)}
          style={{border:"1px solid #E0E0E0",borderRadius:7,padding:"4px 8px",fontSize:11,background:"#fff"}}>
          <option value="temiz">✓ Yok</option><option value="boykot">⛔ Boykot</option><option value="bilinmiyor">? Bilinmiyor</option>
        </select>
      </Row>

      {/* Hamilelik */}
      <Row label="Hamile uyumu">
        <select value={item.pregnancy_safe||"unknown"} onChange={e=>onUpdate(item.id,"pregnancy_safe",e.target.value)}
          style={{border:"1px solid #E0E0E0",borderRadius:7,padding:"4px 8px",fontSize:11,background:"#fff"}}>
          <option value="safe">✓ Güvenli</option><option value="avoid">⛔ Kaçınılmalı</option>
          <option value="consult">⚠ Doktora Danış</option><option value="unknown">? Bilinmiyor</option>
        </select>
      </Row>

      {/* SKT */}
      <Row label="SKT (son kullanma)">
        <input type="date" defaultValue={item.skt||""} onChange={e=>onUpdate(item.id,"skt",e.target.value||null)}
          style={{border:"1px solid #E0E0E0",borderRadius:7,padding:"4px 8px",fontSize:11,color:"#333"}}/>
      </Row>

      {/* Cilt tipi uyumu */}
      <div style={{padding:"8px 0",borderTop:"1px solid #F5F5F5"}}>
        <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:6}}>Cilt Tipi Uyumu (algoritma için)</div>
        <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
          {Object.entries(SKIN_TYPE_LABELS).map(([k,v])=>{
            const has=(item.skin_types||[]).includes(k);
            return <button key={k} onClick={()=>toggleSkinType(k)}
              style={{padding:"3px 9px",borderRadius:6,fontSize:10,fontWeight:600,cursor:"pointer",
                border:`1.5px solid ${has?"#2C4A3E":"#E0E0E0"}`,
                background:has?"#E9F7EF":"transparent",color:has?"#2C4A3E":"#999"}}>{v}</button>;
          })}
        </div>
      </div>

      {/* Endişe etiketleri */}
      <div style={{padding:"8px 0",borderTop:"1px solid #F5F5F5"}}>
        <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:6}}>Endişe Etiketleri (algoritma için)</div>
        <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
          {Object.entries(CONCERN_LABELS).map(([k,v])=>{
            const has=(item.concerns||[]).includes(k);
            return <button key={k} onClick={()=>toggleConcern(k)}
              style={{padding:"3px 9px",borderRadius:6,fontSize:10,fontWeight:600,cursor:"pointer",
                border:`1.5px solid ${has?"#C4974A":"#E0E0E0"}`,
                background:has?"#FEF3E0":"transparent",color:has?"#7D4700":"#999"}}>{v}</button>;
          })}
        </div>
      </div>

      {/* Fiyat + Maliyet */}
      <div style={{padding:"8px 0",borderTop:"1px solid #F5F5F5"}}>
        <div style={{display:"flex",gap:8}}>
          <div style={{flex:1}}>
            <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:4}}>Satış Fiyatı (₺)</div>
            <input type="number" value={price} onChange={e=>setPrice(e.target.value)}
              style={{width:"100%",padding:"7px 10px",border:"1px solid #E0E0E0",borderRadius:7,
                fontSize:13,fontWeight:600,color:"#2C4A3E",outline:"none",boxSizing:"border-box"}}/>
          </div>
          <div style={{flex:1}}>
            <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:4}}>Maliyet (₺)</div>
            <input type="number" value={cost} onChange={e=>setCost(e.target.value)} placeholder="—"
              style={{width:"100%",padding:"7px 10px",border:"1px solid #E0E0E0",borderRadius:7,
                fontSize:13,color:"#333",outline:"none",boxSizing:"border-box"}}/>
          </div>
        </div>
      </div>

      {/* Fotoğraf: sadece süper admin (Havuz) tarafından yönetilir */}
      {!item.photo&&<div style={{padding:"6px 0",fontSize:10,color:"#AAA",fontStyle:"italic"}}>
        📷 Fotoğraf bu üründe tanımlı değil — Havuz'dan süper admin tarafından eklenir.
      </div>}

      {/* Manuel indirim */}
      <div style={{padding:"8px 0",borderTop:"1px solid #F5F5F5"}}>
        <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:6}}>Manuel İndirim</div>
        <div style={{display:"flex",gap:5}}>
          {[5,10,15,20].map(pct=>(
            <button key={pct} onClick={()=>{
              const base=parseInt(price)||(item.price||item.basePrice);
              const c=cost?parseInt(cost):null;
              if(c){const d=base*(1-pct/100);if(d<c*(1+MIN_MARGIN)){alert("⛔ Min. kâr marjı ihlali!");return;}}
              onUpdate(item.id,"discountedPrice",Math.round(base*(1-pct/100)));
            }} style={{flex:1,background:"#F0F4F0",color:"#2C4A3E",border:"1px solid #2C4A3E",
              borderRadius:7,padding:"6px 0",fontSize:11,fontWeight:600,cursor:"pointer"}}>%{pct}</button>
          ))}
          <button onClick={()=>onUpdate(item.id,"discountedPrice",null)}
            style={{flex:1,background:"#FDECEA",color:"#C0392B",border:"none",
              borderRadius:7,padding:"6px 0",fontSize:11,fontWeight:600,cursor:"pointer"}}>Sıfırla</button>
        </div>
      </div>

      {/* Sertifikalar */}
      <div style={{padding:"8px 0",borderTop:"1px solid #F5F5F5"}}>
        <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:6}}>Sertifikalar</div>
        <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
          {Object.entries(CERT_META).map(([cid,cm])=>{
            const has=(item.certs||[]).includes(cid);
            return <button key={cid}
              onClick={()=>onUpdate(item.id,"certs",has?(item.certs||[]).filter(c=>c!==cid):[...(item.certs||[]),cid])}
              style={{padding:"3px 9px",borderRadius:6,fontSize:10,fontWeight:600,cursor:"pointer",
                border:`1.5px solid ${has?cm.color:"#E0E0E0"}`,background:has?cm.bg:"transparent",color:has?cm.color:"#999"}}>
              {cm.label}</button>;
          })}
        </div>
      </div>

      <button onClick={save}
        style={{width:"100%",marginTop:10,background:saved?"#E9F7EF":"#1C1C1A",
          color:saved?"#1E6B3E":"#fff",border:saved?"1px solid #C3E6CB":"none",
          borderRadius:10,padding:"10px",fontSize:13,fontWeight:700,cursor:"pointer",transition:"all .2s"}}>
        {saved?"✓ Kaydedildi":"💾 Kaydet"}
      </button>
    </div>
  );
}

// Yardımcı Row bileşeni
function Row({label,children}){
  return <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",
    padding:"7px 0",borderTop:"1px solid #F5F5F5"}}>
    <span style={{fontSize:12,fontWeight:500,color:"#333"}}>{label}</span>
    {children}
  </div>;
}

// ══════════════════════════════════════════════════════════════
// MARKA / KATEGORİ FİLTRE ÇİPLERİ (admin listeleri için)
// ══════════════════════════════════════════════════════════════
function FilterChips({label,options,value,onChange,allLabel="Tümü",getLabel}){
  if(!options||options.length===0)return null;
  const showLabel=opt=>getLabel?getLabel(opt):opt;
  return(
    <div style={{marginBottom:8}}>
      {label&&<div style={{fontSize:10,fontWeight:600,color:"#999",marginBottom:4}}>{label}</div>}
      <div style={{display:"flex",gap:6,overflowX:"auto",paddingBottom:2,WebkitOverflowScrolling:"touch"}}>
        {["",...options].map(opt=>(
          <button key={opt||"__all__"} onClick={()=>onChange(opt)}
            style={{flexShrink:0,padding:"6px 12px",borderRadius:20,fontSize:11,fontWeight:600,cursor:"pointer",
              border:`1.5px solid ${value===opt?"#1C1C1A":"#E0E0E0"}`,
              background:value===opt?"#1C1C1A":"#fff",color:value===opt?"#fff":"#555",whiteSpace:"nowrap"}}>
            {opt===""?allLabel:showLabel(opt)}
          </button>
        ))}
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════
// KPI KARTI
// ══════════════════════════════════════════════════════════════
function KPI({label,value,sub,color="#1C1C1A"}){
  return(
    <div style={{background:"#fff",borderRadius:12,border:"1px solid #E0E0E0",padding:12}}>
      <div style={{fontSize:10,color:"#888",marginBottom:4}}>{label}</div>
      <div style={{fontSize:16,fontWeight:700,color}}>{value}</div>
      {sub&&<div style={{fontSize:10,color:"#999",marginTop:2}}>{sub}</div>}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════
// SATIŞ GRAFİĞİ + FAYDA/ZARAR ANALİZİ (Eczane paneli)
// Not: "sales", kiosk üzerinden "Listeyi Tamamla" ile tamamlanan
// listelerden oluşur (ödeme kasada yapılır). Gerçek bir POS
// entegrasyonu değildir; kiosk kullanım göstergesi olarak okunmalıdır.
// ══════════════════════════════════════════════════════════════
function AnalyticsPanel({sales,pharmacyCatalog,onResetSales}){
  const now=new Date();
  const days=[...Array(14)].map((_,i)=>{
    const d=new Date(now); d.setDate(d.getDate()-(13-i)); d.setHours(0,0,0,0);
    return d;
  });
  const dayTotals=days.map(d=>{
    const next=new Date(d); next.setDate(d.getDate()+1);
    const total=sales.filter(s=>{const sd=new Date(s.date); return sd>=d&&sd<next;})
      .reduce((sum,s)=>sum+(s.total||0),0);
    return {date:d,total};
  });
  const maxVal=Math.max(1,...dayTotals.map(d=>d.total));

  const totalRevenue=sales.reduce((s,x)=>s+(x.total||0),0);
  const totalOrig=sales.reduce((s,x)=>s+(x.totalOrig||0),0);
  const totalSavings=Math.max(0,totalOrig-totalRevenue);
  const salesCount=sales.length;
  const avgBasket=salesCount?Math.round(totalRevenue/salesCount):0;
  const avgItems=salesCount?(sales.reduce((s,x)=>s+(x.itemCount||0),0)/salesCount):0;

  const weekAgo=new Date(now); weekAgo.setDate(now.getDate()-7);
  const twoWeeksAgo=new Date(now); twoWeeksAgo.setDate(now.getDate()-14);
  const thisWeekTotal=sales.filter(s=>new Date(s.date)>=weekAgo).reduce((s,x)=>s+(x.total||0),0);
  const lastWeekTotal=sales.filter(s=>{const d=new Date(s.date);return d>=twoWeeksAgo&&d<weekAgo;}).reduce((s,x)=>s+(x.total||0),0);
  const weekChange=lastWeekTotal>0?Math.round(((thisWeekTotal-lastWeekTotal)/lastWeekTotal)*100):(thisWeekTotal>0?100:0);

  const prodMap=new Map();
  sales.forEach(s=>(s.items||[]).forEach(it=>{
    const cur=prodMap.get(it.id)||{name:it.name,brand:it.brand,qty:0,revenue:0};
    cur.qty+=(it.qty||1); cur.revenue+=(it.unitPrice||0)*(it.qty||1);
    prodMap.set(it.id,cur);
  }));
  const topProducts=[...prodMap.values()].sort((a,b)=>b.qty-a.qty).slice(0,5);

  const discountedSalesCount=sales.filter(s=>s.discountApplied).length;
  const discountedShare=salesCount?Math.round((discountedSalesCount/salesCount)*100):0;

  const urgentItems=pharmacyCatalog.filter(p=>{const d=daysUntil(p.skt);return d!=null&&d>=0&&d<=30;});
  const stagnantItems=pharmacyCatalog.filter(p=>getMovement(p,sales).label==="Hareketsiz");
  const suggestCount=pharmacyCatalog.filter(p=>suggestDiscount(p,sales)).length;

  const marginViolations=pharmacyCatalog.filter(p=>{
    if(!p.discountedPrice||!p.cost)return false;
    const margin=(p.discountedPrice-p.cost)/p.discountedPrice;
    return margin<MIN_MARGIN-0.001;
  }).length;

  const benefits=[]; const concerns=[];
  if(avgItems>1.3)benefits.push(`Ortalama sepet ${avgItems.toFixed(1)} ürün — kiosk çapraz satışı destekliyor.`);
  if(discountedShare>0)benefits.push(`Tamamlanan listelerin %${discountedShare}'i onaylı bir indirimle kapandı — stok erimesine katkı sağlıyor.`);
  if(marginViolations===0&&pharmacyCatalog.some(p=>p.discountedPrice))
    benefits.push(`Uygulanan indirimlerde minimum kâr marjı (%${Math.round(MIN_MARGIN*100)}) hiçbir üründe ihlal edilmedi.`);
  else if(marginViolations>0)
    concerns.push(`${marginViolations} üründe indirimli fiyat, belirlenen minimum kâr marjının altında — kontrol edin.`);
  if(urgentItems.length>0)benefits.push(`${urgentItems.length} ürünün miadı 30 gün içinde doluyor; oto indirim önerisiyle elden çıkarma şansı var.`);
  if(stagnantItems.length>2)concerns.push(`${stagnantItems.length} ürün "hareketsiz" görünüyor — haftalık satış verisini güncelleyin veya indirime alın.`);
  if(salesCount===0)concerns.push("Henüz kiosk üzerinden tamamlanmış bir liste yok; grafik ve göstergeler kullanımla birlikte anlamlanacak.");

  return(
    <div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}>
        <KPI label="Toplam Ciro (kiosk)" value={totalRevenue.toLocaleString("tr-TR")+" ₺"} sub={`${salesCount} liste`}/>
        <KPI label="Ort. Sepet Tutarı" value={avgBasket.toLocaleString("tr-TR")+" ₺"} sub={`${avgItems.toFixed(1)} ürün/liste`}/>
        <KPI label="Bu Hafta vs Geçen Hafta" value={(weekChange>=0?"+":"")+weekChange+"%"}
          sub={thisWeekTotal.toLocaleString("tr-TR")+" ₺ bu hafta"} color={weekChange>=0?"#1E6B3E":"#C0392B"}/>
        <KPI label="Müşteriye Sağlanan İndirim" value={totalSavings.toLocaleString("tr-TR")+" ₺"} sub="toplam"/>
      </div>

      <div style={{background:"#fff",borderRadius:14,border:"1px solid #E0E0E0",padding:14,marginBottom:12}}>
        <div style={{fontSize:12,fontWeight:700,color:"#1C1C1A",marginBottom:10}}>Son 14 Gün — Kiosk Satış Grafiği (₺)</div>
        {salesCount===0
          ?<div style={{textAlign:"center",padding:20,color:"#999",fontSize:12}}>Henüz veri yok.</div>
          :<svg viewBox="0 0 284 110" width="100%" height="110">
            {dayTotals.map((d,i)=>{
              const barH=Math.max(2,(d.total/maxVal)*78);
              const x=i*20+4;
              return <g key={i}>
                <rect x={x} y={92-barH} width={12} height={barH} rx={2} fill="#2C4A3E"/>
                <text x={x+6} y={104} fontSize="6" fill="#999" textAnchor="middle">{d.date.getDate()}</text>
              </g>;
            })}
          </svg>}
      </div>

      {topProducts.length>0&&(
        <div style={{background:"#fff",borderRadius:14,border:"1px solid #E0E0E0",padding:14,marginBottom:12}}>
          <div style={{fontSize:12,fontWeight:700,color:"#1C1C1A",marginBottom:8}}>En Çok Satan Ürünler</div>
          {topProducts.map((p,i)=>(
            <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",
              borderBottom:i<topProducts.length-1?"1px solid #F5F5F5":"none"}}>
              <div style={{fontSize:12,color:"#333"}}>{i+1}. {p.brand} {p.name}</div>
              <div style={{fontSize:12,fontWeight:700,color:"#2C4A3E"}}>{p.qty} adet</div>
            </div>
          ))}
        </div>
      )}

      <div style={{background:"#fff",borderRadius:14,border:"1px solid #E0E0E0",padding:14,marginBottom:12}}>
        <div style={{fontSize:12,fontWeight:700,color:"#1C1C1A",marginBottom:8}}>Oto İndirim / Stok Özeti</div>
        <div style={{fontSize:12,color:"#555",lineHeight:1.8}}>
          <div>💡 {suggestCount} üründe indirim önerisi var</div>
          <div>⏳ {urgentItems.length} ürünün miadı 30 gün içinde doluyor</div>
          <div>📉 {stagnantItems.length} ürün hareketsiz görünüyor</div>
        </div>
      </div>

      <div style={{background:"#EEF0FB",borderRadius:14,padding:16,marginBottom:12}}>
        <div style={{fontSize:13,fontWeight:700,color:"#1A237E",marginBottom:8}}>📊 Sistemin Eczaneye Katkısı</div>
        {benefits.length>0&&<div style={{marginBottom:concerns.length>0?10:0}}>
          {benefits.map((b,i)=><div key={i} style={{fontSize:12,color:"#1E6B3E",marginBottom:5,lineHeight:1.5}}>✓ {b}</div>)}
        </div>}
        {concerns.length>0&&<div>
          {concerns.map((c,i)=><div key={i} style={{fontSize:12,color:"#8B2E2E",marginBottom:5,lineHeight:1.5}}>⚠ {c}</div>)}
        </div>}
        <div style={{fontSize:10,color:"#666",marginTop:8,lineHeight:1.5,fontStyle:"italic"}}>
          Bu bölüm, kiosk kullanım verisine dayanan kural tabanlı bir göstergedir; kesin bir mali/ROI hesabı değildir.
          Sağlıklı bir değerlendirme için gerçek kasa (POS) verileriyle karşılaştırmanız önerilir.
        </div>
      </div>

      {onResetSales&&(
        <div style={{background:"#fff",borderRadius:14,border:"1px solid #E0E0E0",padding:14,marginBottom:12}}>
          <div style={{fontSize:12,fontWeight:700,color:"#1C1C1A",marginBottom:4}}>Analiz Verilerini Sıfırla</div>
          <div style={{fontSize:11,color:"#888",marginBottom:10,lineHeight:1.5}}>
            Kayıtlı tüm kiosk satış geçmişini siler; katalog ve ürün bilgileri etkilenmez. Bu işlem geri alınamaz.
          </div>
          <button onClick={()=>{
              if(salesCount===0)return;
              if(window.confirm(`${salesCount} satış kaydı kalıcı olarak silinecek. Emin misiniz?`))onResetSales();
            }}
            disabled={salesCount===0}
            style={{width:"100%",background:salesCount===0?"#F0F0F0":"#FDECEA",
              color:salesCount===0?"#AAA":"#C0392B",border:"none",borderRadius:10,
              padding:"10px",fontSize:12,fontWeight:600,cursor:salesCount===0?"default":"pointer"}}>
            🗑 Analiz Verilerini Sıfırla
          </button>
        </div>
      )}
    </div>
  );
}

function PharmacyAdmin({masterPool,setPool,pharmacyCatalog,setPharmacyCatalog,pharmacySettings,setPharmacySettings,catalogReady=true,sales=[],onResetSales,account,onUpdateAccount,onBack}){
  const T=useTheme(pharmacySettings);
  const [tab,setTab]=useState("catalog");
  const autoNavRef=useRef(false);
  useEffect(()=>{
    if(catalogReady&&!autoNavRef.current){
      autoNavRef.current=true;
      if(pharmacyCatalog.length===0)setTab("import");
    }
  },[catalogReady]); // eslint-disable-line
  const [linkCopied,setLinkCopied]=useState(false);
  const kioskLink=typeof window!=="undefined"
    ?`${window.location.origin}${window.location.pathname}?store=${account?.id||""}`
    :"";
  // Güvenlik: 10 dakika etkileşim olmazsa panelden otomatik çıkış yapılır
  useIdleTimer(10*60*1000,onBack,true);
  const [search,setSearch]=useState("");
  const [filterBrand,setFilterBrand]=useState("");
  const [filterCat,setFilterCat]=useState("");
  const [sortBy,setSortBy]=useState("name");
  const [selected,setSelected]=useState(new Set());
  const [bulkStock,setBulkStock]=useState("");
  const [msg,setMsg]=useState("");
  const [showScanner,setShowScanner]=useState(false);
  const [showCredModal,setShowCredModal]=useState(false);

  const switchTab=id=>{setSearch("");setFilterBrand("");setFilterCat("");setSelected(new Set());setTab(id);};
  const toggleSelect=id=>setSelected(prev=>{const n=new Set(prev);n.has(id)?n.delete(id):n.add(id);return n;});

  const inCatalogIds=new Set(pharmacyCatalog.map(p=>p.id));
  const notify=m=>{setMsg(m);setTimeout(()=>setMsg(""),2500);};

  const addFromPool=p=>{
    if(inCatalogIds.has(p.id))return;
    setPharmacyCatalog(prev=>[...prev,{...p,price:p.basePrice,stock:true,cost:null,discountedPrice:null}]);
    notify(`✓ "${p.name}" kataloğa eklendi.`);
  };
  const removeFromCatalog=id=>{setPharmacyCatalog(prev=>prev.filter(p=>p.id!==id));notify("Ürün çıkarıldı.");};
  const updateCatalogItem=(id,f,v)=>setPharmacyCatalog(prev=>prev.map(p=>p.id===id?{...p,[f]:v}:p));

  const handleScanFound=p=>{
    if(inCatalogIds.has(p.id)){notify(`"${p.name}" zaten kataloğunuzda.`);return;}
    addFromPool(p);
  };

  const catalogBrands=[...new Set(pharmacyCatalog.map(p=>p.brand))].sort((a,b)=>a.localeCompare(b,"tr"));
  const poolBrands=[...new Set(masterPool.filter(p=>!inCatalogIds.has(p.id)).map(p=>p.brand))].sort((a,b)=>a.localeCompare(b,"tr"));

  const sortFn=(a,b)=>{
    if(sortBy==="name")return a.name.localeCompare(b.name,"tr");
    if(sortBy==="price_asc")return(a.price||a.basePrice)-(b.price||b.basePrice);
    if(sortBy==="price_desc")return(b.price||b.basePrice)-(a.price||a.basePrice);
    if(sortBy==="stock_asc"){const sa=stockQty(a),sb=stockQty(b);return(sa??Infinity)-(sb??Infinity);}
    return 0;
  };

  const filteredCatalog=pharmacyCatalog
    .filter(p=>filterBrand===""||p.brand===filterBrand)
    .filter(p=>filterCat===""||p.category===filterCat)
    .filter(p=>search===""||
      p.name.toLowerCase().includes(search.toLowerCase())||p.brand.toLowerCase().includes(search.toLowerCase()))
    .sort(sortFn);
  const filteredPool=masterPool
    .filter(p=>!inCatalogIds.has(p.id))
    .filter(p=>filterBrand===""||p.brand===filterBrand)
    .filter(p=>filterCat===""||p.category===filterCat)
    .filter(p=>search===""||
      p.name.toLowerCase().includes(search.toLowerCase())||p.brand.toLowerCase().includes(search.toLowerCase()))
    .sort(sortFn);

  // Katalog özet istatistikleri
  const lowStockCount=pharmacyCatalog.filter(p=>{const q=stockQty(p);return q!=null&&q>0&&q<10;}).length;
  const outStockCount=pharmacyCatalog.filter(p=>!hasStock(p)).length;
  const expiringCount=pharmacyCatalog.filter(p=>{const d=daysUntil(p.skt);return d!=null&&d>=0&&d<=30;}).length;

  // Algoritmanın indirim önerdiği ürünler (aksiyon listesi)
  const suggestedItems=pharmacyCatalog
    .map(p=>({p,sug:suggestDiscount(p,sales)}))
    .filter(x=>x.sug);

  // Toplu işlemler
  const bulkDeleteCatalog=()=>{
    if(selected.size===0)return;
    if(!window.confirm(`${selected.size} ürün kataloğunuzdan çıkarılsın mı?`))return;
    setPharmacyCatalog(prev=>prev.filter(p=>!selected.has(p.id)));
    notify(`${selected.size} ürün çıkarıldı.`);
    setSelected(new Set());
  };
  const bulkSetStock=()=>{
    const v=parseInt(bulkStock);
    if(isNaN(v)||v<0||selected.size===0)return;
    setPharmacyCatalog(prev=>prev.map(p=>selected.has(p.id)?{...p,stock:v}:p));
    notify(`${selected.size} ürünün stoğu ${v} olarak güncellendi.`);
    setSelected(new Set());setBulkStock("");
  };
  const applyAllSuggestions=()=>{
    if(suggestedItems.length===0)return;
    if(!window.confirm(`${suggestedItems.length} üründe algoritmanın önerdiği indirim uygulansın mı?`))return;
    const ids=new Set(suggestedItems.map(x=>x.p.id));
    setPharmacyCatalog(prev=>prev.map(p=>{
      if(!ids.has(p.id))return p;
      const sug=suggestedItems.find(x=>x.p.id===p.id).sug;
      const base=p.price||p.basePrice;
      return{...p,discountedPrice:Math.round(base*(1-sug.percent/100))};
    }));
    notify(`✓ ${suggestedItems.length} üründe otomatik indirim uygulandı.`);
  };

  return(
    <div className="admin-shell" style={{background:T.bg,minHeight:"100vh",
      paddingBottom:80,fontFamily:"'Inter',sans-serif"}}>
      {showScanner&&<BarcodeScanner catalog={masterPool} onFound={handleScanFound}
        onClose={()=>setShowScanner(false)} mode="pharmacy"/>}
      {showCredModal&&<ChangeCredentialsModal title="Hesap Bilgileri"
        initialEmail={account?.email} initialPhone={account?.phone} showPhone={true}
        onSave={fields=>onUpdateAccount(account.id,fields)} onClose={()=>setShowCredModal(false)}/>}

      <div style={{background:"#1C1C1A",color:"#fff",padding:"14px 16px 0",position:"sticky",top:0,zIndex:50}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
          <div>
            <div style={{fontSize:16,fontWeight:700}}>🏪 Eczane Yönetimi</div>
            <div style={{fontSize:11,opacity:.5,marginTop:1}}>{pharmacySettings?.name||"Eczane"} · {pharmacyCatalog.length} ürün</div>
            {account&&<div style={{fontSize:10,opacity:.4,marginTop:1}}>👤 {account.contactName} · {account.email}</div>}
          </div>
          <div style={{display:"flex",gap:7}}>
            <button onClick={()=>setShowScanner(true)}
              style={{background:"rgba(255,255,255,.1)",border:"none",color:"#fff",
                borderRadius:8,padding:"6px 10px",fontSize:13,cursor:"pointer"}}>📷 Tara</button>
            <button onClick={onBack} style={{background:"rgba(255,255,255,.1)",border:"none",color:"#fff",
              borderRadius:8,padding:"6px 10px",fontSize:12,cursor:"pointer"}}>← Çıkış</button>
          </div>
        </div>
        <div style={{display:"flex",gap:4,paddingBottom:0}}>
          {[{id:"catalog",l:"Katalogum"},{id:"pool",l:"Havuz"},{id:"import",l:"İçe Aktar"},{id:"analiz",l:"Analiz"},{id:"settings",l:"Ayarlar"}].map(t=>(
            <button key={t.id} onClick={()=>switchTab(t.id)}
              style={{flex:1,background:tab===t.id?"rgba(255,255,255,.15)":"transparent",border:"none",
                color:tab===t.id?"#fff":"rgba(255,255,255,.5)",borderRadius:"8px 8px 0 0",
                padding:"7px 4px",fontSize:11,fontWeight:tab===t.id?700:400,cursor:"pointer"}}>
              {t.l}
            </button>
          ))}
        </div>
      </div>

      <div style={{padding:"12px 14px 0"}}>
        <Toast msg={msg}/>
        {(tab==="catalog"||tab==="pool")&&(<>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Ara…"
            style={{width:"100%",padding:"10px 14px",border:"1.5px solid #E0E0E0",borderRadius:12,
              fontSize:14,background:"#fff",outline:"none",boxSizing:"border-box",marginBottom:10}}/>
          <FilterChips label="Marka" options={tab==="catalog"?catalogBrands:poolBrands}
            value={filterBrand} onChange={setFilterBrand}/>
          <FilterChips label="Kategori" options={Object.keys(CAT_LABELS)}
            value={filterCat} onChange={setFilterCat}
            getLabel={c=>CAT_LABELS[c]} allLabel="Tümü"/>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
            <div style={{fontSize:11,color:"#999"}}>
              {tab==="catalog"?filteredCatalog.length:filteredPool.length} ürün
            </div>
            <select value={sortBy} onChange={e=>setSortBy(e.target.value)}
              style={{padding:"6px 8px",border:"1px solid #E0E0E0",borderRadius:8,fontSize:11,
                background:"#fff",color:"#555"}}>
              <option value="name">Ada göre (A-Z)</option>
              <option value="price_asc">Fiyata göre (artan)</option>
              <option value="price_desc">Fiyata göre (azalan)</option>
              {tab==="catalog"&&<option value="stock_asc">Stoğa göre (az→çok)</option>}
            </select>
          </div>
        </>)}
        {tab==="catalog"&&catalogReady&&pharmacyCatalog.length>0&&(
          <div style={{display:"flex",gap:6,marginBottom:10,overflowX:"auto"}}>
            <div style={{flex:1,minWidth:80,background:"#fff",borderRadius:10,border:"1px solid #E0E0E0",
              padding:"8px 10px",textAlign:"center"}}>
              <div style={{fontSize:16,fontWeight:700,color:"#1C1C1A"}}>{pharmacyCatalog.length}</div>
              <div style={{fontSize:9,color:"#999"}}>Ürün</div>
            </div>
            <div style={{flex:1,minWidth:80,background:outStockCount>0?"#FDECEA":"#fff",borderRadius:10,
              border:`1px solid ${outStockCount>0?"#F5C6C0":"#E0E0E0"}`,padding:"8px 10px",textAlign:"center"}}>
              <div style={{fontSize:16,fontWeight:700,color:outStockCount>0?"#8B2E2E":"#1C1C1A"}}>{outStockCount}</div>
              <div style={{fontSize:9,color:outStockCount>0?"#8B2E2E":"#999"}}>Stok Yok</div>
            </div>
            <div style={{flex:1,minWidth:80,background:lowStockCount>0?"#FFF3E0":"#fff",borderRadius:10,
              border:`1px solid ${lowStockCount>0?"#F5DBA8":"#E0E0E0"}`,padding:"8px 10px",textAlign:"center"}}>
              <div style={{fontSize:16,fontWeight:700,color:lowStockCount>0?"#7D4700":"#1C1C1A"}}>{lowStockCount}</div>
              <div style={{fontSize:9,color:lowStockCount>0?"#7D4700":"#999"}}>Az Stok</div>
            </div>
            <div style={{flex:1,minWidth:80,background:expiringCount>0?"#FFF3E0":"#fff",borderRadius:10,
              border:`1px solid ${expiringCount>0?"#F5DBA8":"#E0E0E0"}`,padding:"8px 10px",textAlign:"center"}}>
              <div style={{fontSize:16,fontWeight:700,color:expiringCount>0?"#7D4700":"#1C1C1A"}}>{expiringCount}</div>
              <div style={{fontSize:9,color:expiringCount>0?"#7D4700":"#999"}}>SKT Yakın</div>
            </div>
          </div>
        )}
        {tab==="catalog"&&suggestedItems.length>0&&(
          <div style={{background:"#FFFDE7",border:"1px solid #FBC02D",borderRadius:12,
            padding:"12px 14px",marginBottom:10}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,flexWrap:"wrap"}}>
              <div>
                <div style={{fontSize:13,fontWeight:700,color:"#7D4700"}}>
                  💡 {suggestedItems.length} üründe algoritma indirim öneriyor
                </div>
                <div style={{fontSize:10,color:"#7D4700",opacity:.8,marginTop:2}}>
                  Stok devri, satış hızı, miad ve kârlılığa göre otomatik hesaplandı
                </div>
              </div>
              <button onClick={applyAllSuggestions}
                style={{background:"#7D4700",color:"#fff",border:"none",borderRadius:9,
                  padding:"9px 14px",fontSize:12,fontWeight:700,cursor:"pointer",flexShrink:0}}>
                Tümünü Uygula →
              </button>
            </div>
          </div>
        )}
        {tab==="catalog"&&selected.size>0&&(
          <div style={{background:"#1C1C1A",borderRadius:12,padding:"10px 12px",marginBottom:10,
            display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
            <span style={{color:"#fff",fontSize:12,fontWeight:600}}>{selected.size} seçili</span>
            <input type="number" min="0" placeholder="Stok adedi" value={bulkStock}
              onChange={e=>setBulkStock(e.target.value)}
              style={{width:90,padding:"6px 8px",border:"none",borderRadius:7,fontSize:11}}/>
            <button onClick={bulkSetStock} disabled={bulkStock===""}
              style={{background:bulkStock===""?"#555":"#2C4A3E",color:"#fff",border:"none",borderRadius:7,
                padding:"6px 10px",fontSize:11,fontWeight:600,cursor:bulkStock===""?"default":"pointer"}}>
              Stoğu Ayarla
            </button>
            <button onClick={bulkDeleteCatalog}
              style={{background:"#FDECEA",color:"#C0392B",border:"none",borderRadius:7,
                padding:"6px 10px",fontSize:11,fontWeight:600,cursor:"pointer"}}>🗑 Kaldır</button>
            <button onClick={()=>setSelected(new Set())}
              style={{background:"none",color:"#fff",border:"none",fontSize:11,cursor:"pointer",
                marginLeft:"auto",opacity:.7}}>Seçimi Temizle</button>
          </div>
        )}
        {tab==="import"&&(
          <ExcelImport
            masterPool={masterPool}
            setPool={setPool}
            pharmacyCatalog={pharmacyCatalog}
            onImport={newCatalog=>{
              setPharmacyCatalog(newCatalog);
              notify(`✓ İçe aktarma tamamlandı. Katalog güncellendi.`);
              setTab("catalog");
            }}
          />
        )}
        {tab==="catalog"&&(
          !catalogReady
            ?<div className="admin-grid">{[1,2,3].map(i=><SkeletonCard key={i}/>)}</div>
            :filteredCatalog.length===0
            ?<div style={{textAlign:"center",padding:40,color:"#999"}}>Katalog boş. Havuz veya 📷 Tara ile ekleyin.</div>
            :<div className="admin-grid fade-in">{filteredCatalog.map(item=>(
              <CatalogItem key={item.id} item={item} onUpdate={updateCatalogItem} onRemove={removeFromCatalog}
                selected={selected.has(item.id)} onToggleSelect={()=>toggleSelect(item.id)} sales={sales}/>
            ))}</div>
        )}
        {tab==="pool"&&(
          <div>
            <div style={{fontSize:11,color:"#999",marginBottom:8}}>Havuzdan seçip kataloğunuza ekleyin</div>
            {filteredPool.length===0&&<div style={{textAlign:"center",padding:40,color:"#999"}}>Tüm ürünler kataloğunuzda.</div>}
            <div className="admin-grid">
            {filteredPool.map(p=>(
              <div key={p.id} style={{background:"#fff",borderRadius:12,border:"1px solid #E0E0E0",
                padding:12,display:"flex",gap:10,alignItems:"center"}}>
                <Avatar brand={p.brand} size={38}/>
                <div style={{flex:1}}>
                  <div style={{fontSize:10,color:"#999"}}>{p.brand}</div>
                  <div style={{fontSize:13,fontWeight:600,color:"#1C1C1A"}}>{p.name}</div>
                  <div style={{fontSize:11,fontWeight:700,color:"#2C4A3E",marginTop:1}}>
                    {p.basePrice.toLocaleString("tr-TR")} ₺ · {CAT_LABELS[p.category]}
                  </div>
                  <div style={{display:"flex",flexWrap:"wrap",gap:3,marginTop:4}}>
                    {(p.skin_types||[]).map(st=><span key={st} style={{background:"#E9F7EF",color:"#2C4A3E",fontSize:9,padding:"1px 5px",borderRadius:4}}>{SKIN_TYPE_LABELS[st]}</span>)}
                    {(p.concerns||[]).map(c=><span key={c} style={{background:"#FEF3E0",color:"#7D4700",fontSize:9,padding:"1px 5px",borderRadius:4}}>{CONCERN_LABELS[c]}</span>)}
                    {p.pregnancy_safe==="safe"&&<span style={{background:"#E9F7EF",color:"#1E6B3E",fontSize:9,padding:"1px 5px",borderRadius:4}}>🤰 Güvenli</span>}
                    {p.pregnancy_safe==="avoid"&&<span style={{background:"#FDECEA",color:"#8B2E2E",fontSize:9,padding:"1px 5px",borderRadius:4}}>🤰 Kaçın</span>}
                  </div>
                </div>
                <button onClick={()=>addFromPool(p)}
                  style={{background:"#2C4A3E",color:"#fff",border:"none",borderRadius:9,
                    padding:"8px 14px",fontSize:12,fontWeight:600,cursor:"pointer",flexShrink:0}}>+ Ekle</button>
              </div>
            ))}
            </div>
          </div>
        )}
        {tab==="analiz"&&(
          <AnalyticsPanel sales={sales} pharmacyCatalog={pharmacyCatalog} onResetSales={onResetSales}/>
        )}
        {tab==="settings"&&(
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            <div style={{background:"#fff",borderRadius:14,border:"1px solid #E0E0E0",padding:16}}>
              <div style={{fontSize:13,fontWeight:700,color:"#1C1C1A",marginBottom:10}}>Eczane Adı</div>
              <input defaultValue={pharmacySettings?.name||""} placeholder="Eczane adını girin"
                onBlur={e=>setPharmacySettings(prev=>({...prev,name:e.target.value}))}
                style={{width:"100%",padding:"9px 12px",border:"1px solid #E0E0E0",borderRadius:8,
                  fontSize:13,outline:"none",boxSizing:"border-box"}}/>
            </div>
            <div style={{background:"#fff",borderRadius:14,border:"1px solid #E0E0E0",padding:16}}>
              <div style={{fontSize:13,fontWeight:700,color:"#1C1C1A",marginBottom:4}}>Renk Teması</div>
              <div style={{fontSize:11,color:"#999",marginBottom:10}}>Dominant marka seçin.</div>
              <div style={{display:"flex",flexDirection:"column",gap:6}}>
                {Object.entries(BRAND_THEMES).map(([brand,theme])=>{
                  const sel=(pharmacySettings?.primaryBrand||"Nötr")===brand;
                  return <button key={brand} onClick={()=>setPharmacySettings(prev=>({...prev,primaryBrand:brand}))}
                    style={{display:"flex",alignItems:"center",gap:10,padding:"10px 12px",borderRadius:10,
                      border:`2px solid ${sel?theme.primary:"#E0E0E0"}`,
                      background:sel?theme.bg:"#fff",cursor:"pointer",textAlign:"left"}}>
                    <div style={{width:20,height:20,borderRadius:5,background:theme.primary,flexShrink:0}}/>
                    <div>
                      <div style={{fontSize:12,fontWeight:600,color:"#1C1C1A"}}>{brand}</div>
                      <div style={{fontSize:10,color:"#999"}}>{theme.name}</div>
                    </div>
                    {sel&&<span style={{marginLeft:"auto",color:theme.primary}}>✓</span>}
                  </button>;
                })}
              </div>
            </div>
            {/* Sepet İndirimi Ayarları */}
            <div style={{background:"#fff",borderRadius:14,border:"1px solid #E0E0E0",padding:16}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
                <div style={{fontSize:13,fontWeight:700,color:"#1C1C1A"}}>🎁 Sepet İndirimi</div>
                <Sw value={pharmacySettings?.cartDiscountEnabled!==false}
                  onChange={v=>setPharmacySettings(prev=>({...prev,cartDiscountEnabled:v}))} color="#F9A825"/>
              </div>
              <div style={{fontSize:11,color:"#999",marginBottom:12,lineHeight:1.5}}>
                Müşteri sepete belirlediğiniz sayıda ürün ekleyince, tüm sepet fiyatı otomatik
                olarak indirimli gösterilir (kâr marjı güvenlik sınırının altına düşülmez).
              </div>
              <div style={{display:"flex",gap:8,marginBottom:4}}>
                <div style={{flex:1}}>
                  <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:4}}>Kaç üründen itibaren</div>
                  <input type="number" min="1" defaultValue={pharmacySettings?.cartDiscountThreshold||1}
                    onBlur={e=>{const v=Math.max(1,parseInt(e.target.value)||1);
                      setPharmacySettings(prev=>({...prev,cartDiscountThreshold:v}));}}
                    style={{width:"100%",padding:"8px 10px",border:"1px solid #E0E0E0",borderRadius:8,
                      fontSize:13,outline:"none",boxSizing:"border-box"}}/>
                </div>
                <div style={{flex:1}}>
                  <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:4}}>İndirim Oranı (%)</div>
                  <input type="number" min="1" max="90" defaultValue={pharmacySettings?.cartDiscountPct||10}
                    onBlur={e=>{const v=Math.max(1,Math.min(90,parseInt(e.target.value)||10));
                      setPharmacySettings(prev=>({...prev,cartDiscountPct:v}));}}
                    style={{width:"100%",padding:"8px 10px",border:"1px solid #E0E0E0",borderRadius:8,
                      fontSize:13,outline:"none",boxSizing:"border-box"}}/>
                </div>
              </div>
              <div style={{fontSize:10,color:"#AAA",marginTop:8,lineHeight:1.5}}>
                Örn: eşik = 1 ise sepete ilk ürün eklenir eklenmez indirim uygulanır. Fiyatlar hem sepette
                hem ürün kartlarında üstü çizili orijinal fiyat + indirimli fiyat olarak gösterilir.
              </div>
            </div>
            {/* Kiosk Linki */}
            <div style={{background:"#fff",borderRadius:14,border:"1px solid #E0E0E0",padding:16}}>
              <div style={{fontSize:13,fontWeight:700,color:"#1C1C1A",marginBottom:4}}>📱 Kiosk Linkiniz</div>
              <div style={{fontSize:11,color:"#888",marginBottom:10,lineHeight:1.5}}>
                Bu linki eczanenizdeki kiosk cihazına/tablete kaydedin. Giriş gerektirmez, doğrudan
                sizin kataloğunuzu müşterilere gösterir.
              </div>
              <div style={{background:"#F8F8F8",borderRadius:8,padding:"9px 10px",fontSize:11,
                color:"#333",wordBreak:"break-all",fontFamily:"monospace",marginBottom:8}}>
                {kioskLink}
              </div>
              <button onClick={()=>{
                  navigator.clipboard?.writeText(kioskLink);
                  setLinkCopied(true);setTimeout(()=>setLinkCopied(false),2000);
                }}
                style={{width:"100%",background:linkCopied?"#1E6B3E":"#1C1C1A",color:"#fff",border:"none",
                  borderRadius:10,padding:"11px",fontSize:13,fontWeight:600,cursor:"pointer"}}>
                {linkCopied?"✓ Kopyalandı":"🔗 Linki Kopyala"}
              </button>
            </div>
            {/* Hesap Bilgileri */}
            <div style={{background:"#fff",borderRadius:14,border:"1px solid #E0E0E0",padding:16}}>
              <div style={{fontSize:13,fontWeight:700,color:"#1C1C1A",marginBottom:10}}>Hesap Bilgileri</div>
              <div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:12}}>
                <div style={{fontSize:12,color:"#555"}}><b>Eczane:</b> {account?.pharmacyName||"—"}</div>
                <div style={{fontSize:12,color:"#555"}}><b>Yetkili:</b> {account?.contactName||"—"}</div>
                <div style={{fontSize:12,color:"#555"}}><b>E-posta:</b> {account?.email||"—"}</div>
                <div style={{fontSize:12,color:"#555"}}><b>Telefon:</b> {account?.phone||"—"}</div>
              </div>
              <div style={{fontSize:12,color:"#888",marginBottom:12}}>E-posta, telefon veya şifrenizi güncelleyin.</div>
              <button onClick={()=>setShowCredModal(true)}
                style={{width:"100%",background:"#1C1C1A",color:"#fff",border:"none",borderRadius:10,
                  padding:"11px",fontSize:13,fontWeight:600,cursor:"pointer"}}>🔑 Bilgileri Güncelle</button>
            </div>
          </div>
        )}
      </div>

      <div className="admin-shell" style={{position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",width:"100%",
        background:"#fff",borderTop:"1px solid #E0E0E0",display:"flex"}}>
        {[{id:"catalog",l:"Katalogum",i:"📋"},{id:"pool",l:"Havuz",i:"📦"},{id:"import",l:"İçe Aktar",i:"📊"},{id:"analiz",l:"Analiz",i:"📈"},{id:"settings",l:"Ayarlar",i:"⚙️"}].map(t=>(
          <button key={t.id} onClick={()=>switchTab(t.id)}
            style={{flex:1,border:"none",background:"none",padding:"11px 0 9px",cursor:"pointer",
              color:tab===t.id?"#1C1C1A":"#999",fontWeight:tab===t.id?700:400,fontSize:11,
              borderTop:`2.5px solid ${tab===t.id?"#1C1C1A":"transparent"}`,
              display:"flex",flexDirection:"column",alignItems:"center",gap:2}}>
            <span style={{fontSize:17}}>{t.i}</span>{t.l}
          </button>
        ))}
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════
// HAVUZ ÜRÜNİ DÜZENLEYİCİ (Süper Admin — tüm algoritma alanları)
// ══════════════════════════════════════════════════════════════
function PoolItem({product:p, onUpdate, onRemove, selected=false, onToggleSelect}){
  const [open,setOpen]=useState(false);
  const [barcode,setBarcode]=useState(p.barcode||"");
  const [photo,setPhoto]=useState(p.photo||"");
  const [activesText,setActivesText]=useState((p.actives||[]).join(", "));
  const [savedMeta,setSavedMeta]=useState(false);
  const hasAlgo=(p.skin_types?.length>0)||(p.concerns?.length>0);

  const toggleST=(st)=>{
    const cur=p.skin_types||[];
    onUpdate(p.id,"skin_types",cur.includes(st)?cur.filter(x=>x!==st):[...cur,st]);
  };
  const toggleC=(c)=>{
    const cur=p.concerns||[];
    onUpdate(p.id,"concerns",cur.includes(c)?cur.filter(x=>x!==c):[...cur,c]);
  };
  const saveMeta=()=>{
    onUpdate(p.id,"barcode",barcode.trim());
    onUpdate(p.id,"photo",photo.trim()||null);
    onUpdate(p.id,"actives",activesText.split(",").map(s=>s.trim()).filter(Boolean));
    setSavedMeta(true);setTimeout(()=>setSavedMeta(false),2000);
  };
  const applyAnalysis=(result)=>{
    onUpdate(p.id,"skin_types",result.skin_types);
    onUpdate(p.id,"concerns",result.concerns);
    if(result.pregnancy_safe!=="unknown")onUpdate(p.id,"pregnancy_safe",result.pregnancy_safe);
  };

  return(
    <div style={{background:"#fff",borderRadius:12,
      border:`1px solid ${selected?"#1C1C1A":hasAlgo?"#E0E0E0":"#FBC02D"}`,padding:12,
      boxShadow:selected?"0 0 0 2px rgba(28,28,26,.08)":"none"}}>
      {/* Başlık satırı */}
      <div style={{display:"flex",gap:10,alignItems:"center",marginBottom:8}}>
        {onToggleSelect&&<input type="checkbox" checked={selected} onChange={onToggleSelect}
          style={{width:17,height:17,flexShrink:0,cursor:"pointer",accentColor:"#1C1C1A"}}/>}
        {p.photo?<img src={p.photo} style={{width:38,height:38,borderRadius:8,objectFit:"cover",flexShrink:0}}
          onError={e=>e.target.style.display="none"}/>:<Avatar brand={p.brand} size={38}/>}
        <div style={{flex:1}}>
          <div style={{fontSize:10,color:"#999"}}>{p.brand}</div>
          <div style={{fontSize:13,fontWeight:600,color:"#1C1C1A"}}>{p.name}</div>
          <div style={{fontSize:11,color:"#2C4A3E",fontWeight:700,marginTop:1}}>
            {p.basePrice.toLocaleString("tr-TR")} ₺ · {CAT_LABELS[p.category]}
          </div>
          {p.barcode&&<div style={{fontSize:9,color:"#999",fontFamily:"monospace",marginTop:1}}>📦 {p.barcode}</div>}
          {/* Algoritma etiket özeti */}
          <div style={{display:"flex",flexWrap:"wrap",gap:3,marginTop:4}}>
            {(p.skin_types||[]).map(st=><span key={st} style={{background:"#E9F7EF",color:"#2C4A3E",
              fontSize:9,padding:"1px 5px",borderRadius:4}}>{SKIN_TYPE_LABELS[st]}</span>)}
            {(p.concerns||[]).map(c=><span key={c} style={{background:"#FEF3E0",color:"#7D4700",
              fontSize:9,padding:"1px 5px",borderRadius:4}}>{CONCERN_LABELS[c]}</span>)}
            {p.pregnancy_safe&&p.pregnancy_safe!=="unknown"&&<PregBadge status={p.pregnancy_safe}/>}
            {!hasAlgo&&<span style={{background:"#FFF3E0",color:"#E67E22",fontSize:9,
              fontWeight:700,padding:"1px 5px",borderRadius:4}}>⚠ Algoritma değerleri eksik</span>}
          </div>
        </div>
        <button onClick={()=>setOpen(!open)}
          style={{background:open?"#1C1C1A":"#F0F0F0",color:open?"#fff":"#555",border:"none",
            borderRadius:8,padding:"5px 10px",fontSize:11,cursor:"pointer",flexShrink:0}}>
          {open?"▲ Kapat":"▼ Düzenle"}
        </button>
      </div>

      {/* Genişletilmiş düzenleme */}
      {open&&(
        <div style={{borderTop:"1px solid #F5F5F5",paddingTop:10}}>

          {p.barcode&&<div style={{background:"#FAFAFA",borderRadius:8,padding:"8px 4px",marginBottom:10}}>
            <Barcode value={p.barcode} height={26}/>
          </div>}

          {/* Barkod + Fotoğraf (sadece süper admin düzenleyebilir) */}
          <div style={{padding:"8px 0",borderBottom:"1px solid #F5F5F5"}}>
            <div style={{display:"flex",gap:8,marginBottom:8}}>
              <div style={{flex:1}}>
                <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:4}}>Barkod</div>
                <input type="text" value={barcode} onChange={e=>setBarcode(e.target.value)} placeholder="8690..."
                  style={{width:"100%",padding:"7px 10px",border:"1px solid #E0E0E0",borderRadius:7,
                    fontSize:12,fontFamily:"monospace",outline:"none",boxSizing:"border-box"}}/>
              </div>
            </div>
            <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:4}}>📷 Fotoğraf URL</div>
            <input type="url" value={photo} onChange={e=>setPhoto(e.target.value)} placeholder="https://..."
              style={{width:"100%",padding:"7px 10px",border:"1px solid #E0E0E0",borderRadius:7,
                fontSize:11,outline:"none",boxSizing:"border-box",marginBottom:6}}/>
            {photo&&<img src={photo} style={{width:"100%",height:70,objectFit:"cover",borderRadius:7,marginBottom:6}}
              onError={e=>e.target.style.display="none"} onLoad={e=>e.target.style.display="block"}/>}
            <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:4}}>🧬 Aktif İçerikler (virgülle)</div>
            <input type="text" value={activesText} onChange={e=>setActivesText(e.target.value)}
              placeholder="Hyalüronik Asit, Niasinamid, Retinol"
              style={{width:"100%",padding:"7px 10px",border:"1px solid #E0E0E0",borderRadius:7,
                fontSize:11,outline:"none",boxSizing:"border-box",marginBottom:6}}/>
            <button onClick={saveMeta}
              style={{width:"100%",background:savedMeta?"#1E6B3E":"#1C1C1A",color:"#fff",border:"none",
                borderRadius:7,padding:"8px",fontSize:11,fontWeight:600,cursor:"pointer"}}>
              {savedMeta?"✓ Kaydedildi":"Barkod / Fotoğraf / İçerik Kaydet"}
            </button>
            <IngredientAnalysisBox actives={activesText} onApply={applyAnalysis}/>
          </div>

          {/* Boykot */}
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",
            padding:"7px 0",borderBottom:"1px solid #F5F5F5"}}>
            <span style={{fontSize:12,fontWeight:500,color:"#333"}}>Boykot durumu</span>
            <select value={p.boycott} onChange={e=>onUpdate(p.id,"boycott",e.target.value)}
              style={{border:"1px solid #E0E0E0",borderRadius:7,padding:"4px 8px",fontSize:11,background:"#fff"}}>
              <option value="temiz">✓ Boykot Yok</option>
              <option value="boykot">⛔ Boykot</option>
              <option value="bilinmiyor">? Bilinmiyor</option>
            </select>
          </div>

          {/* Hamile güvenliği */}
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",
            padding:"7px 0",borderBottom:"1px solid #F5F5F5"}}>
            <span style={{fontSize:12,fontWeight:500,color:"#333"}}>🤰 Hamile güvenliği</span>
            <select value={p.pregnancy_safe||"unknown"} onChange={e=>onUpdate(p.id,"pregnancy_safe",e.target.value)}
              style={{border:"1px solid #E0E0E0",borderRadius:7,padding:"4px 8px",fontSize:11,background:"#fff"}}>
              <option value="safe">✓ Güvenli</option>
              <option value="avoid">⛔ Kaçınılmalı</option>
              <option value="consult">⚠ Doktora Danış</option>
              <option value="unknown">? Bilinmiyor</option>
            </select>
          </div>

          {/* Cilt tipi uyumu */}
          <div style={{padding:"8px 0",borderBottom:"1px solid #F5F5F5"}}>
            <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:6}}>
              Cilt Tipi Uyumu
              {(p.skin_types||[]).length===0&&<span style={{color:"#E67E22",marginLeft:4,fontWeight:400}}>— seçilmedi!</span>}
            </div>
            <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
              {Object.entries(SKIN_TYPE_LABELS).map(([k,v])=>{
                const has=(p.skin_types||[]).includes(k);
                return <button key={k} onClick={()=>toggleST(k)}
                  style={{padding:"4px 10px",borderRadius:7,fontSize:11,fontWeight:500,cursor:"pointer",
                    border:`1.5px solid ${has?"#2C4A3E":"#E0E0E0"}`,
                    background:has?"#E9F7EF":"#FAFAFA",color:has?"#2C4A3E":"#888"}}>
                  {has?"✓ ":""}{v}
                </button>;
              })}
            </div>
          </div>

          {/* Endişe etiketleri */}
          <div style={{padding:"8px 0",borderBottom:"1px solid #F5F5F5"}}>
            <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:6}}>
              Endişe Etiketleri
              {(p.concerns||[]).length===0&&<span style={{color:"#E67E22",marginLeft:4,fontWeight:400}}>— seçilmedi!</span>}
            </div>
            <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
              {Object.entries(CONCERN_LABELS).map(([k,v])=>{
                const has=(p.concerns||[]).includes(k);
                return <button key={k} onClick={()=>toggleC(k)}
                  style={{padding:"4px 10px",borderRadius:7,fontSize:11,fontWeight:500,cursor:"pointer",
                    border:`1.5px solid ${has?"#C4974A":"#E0E0E0"}`,
                    background:has?"#FEF3E0":"#FAFAFA",color:has?"#7D4700":"#888"}}>
                  {has?"✓ ":""}{v}
                </button>;
              })}
            </div>
          </div>

          {/* Sil butonu */}
          <button onClick={()=>onRemove(p.id)}
            style={{width:"100%",marginTop:8,background:"#FDECEA",color:"#C0392B",border:"none",
              borderRadius:9,padding:"9px",fontSize:12,fontWeight:600,cursor:"pointer"}}>
            Havuzdan Sil
          </button>
        </div>
      )}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════
// HAVUZA TOPLU EXCEL/CSV İÇE AKTARMA (Süper Admin — Master Havuz)
// Barkod eşleşirse mevcut ürünün sadece fiyatı güncellenir,
// eşleşmezse minimum bilgiyle yeni ürün olarak havuza eklenir.
// ══════════════════════════════════════════════════════════════
function PoolExcelImport({pool,onImport}){
  const [rows,setRows]=useState([]);
  const [headers,setHeaders]=useState([]);
  const [map,setMap]=useState({barcode:"",name:"",price:"",brand:""});
  const [preview,setPreview]=useState([]);
  const [step,setStep]=useState("upload"); // upload|map|preview|done
  const [result,setResult]=useState({added:0,updated:0,skipped:0});
  const [loading,setLoading]=useState(false);
  const [importing,setImporting]=useState(false);
  const fileRef=useRef(null);

  const guessField=(col)=>{
    const c=col.toLowerCase().trim();
    if(/barkod|barcode|ean|upc|kod|code/.test(c))return"barcode";
    if(/ürün|urun|ad|name|product|isim|açıklama|aciklama/.test(c))return"name";
    if(/fiyat|price|tutar|ücret|ucret|satis/.test(c))return"price";
    if(/marka|brand/.test(c))return"brand";
    return"";
  };

  const handleFile=async(file)=>{
    if(!file)return;
    setLoading(true);
    try{
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf,{type:"array"});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const data=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
      if(!data.length){alert("Dosya boş.");setLoading(false);return;}
      const firstRow=data[0].map(String);
      const isHeader=firstRow.some(h=>/[a-zçğışöü]/i.test(h));
      const hdrs=isHeader?firstRow:firstRow.map((_,i)=>`Sütun ${i+1}`);
      const dataRows=isHeader?data.slice(1):data;
      setHeaders(hdrs);
      setRows(dataRows.filter(r=>r.some(c=>String(c).trim())));
      const auto={barcode:"",name:"",price:"",brand:""};
      hdrs.forEach((h,i)=>{const f=guessField(h);if(f&&!auto[f])auto[f]=String(i);});
      setMap(auto);
      setStep("map");
    }catch(e){
      alert("Dosya okunamadı. Lütfen .xlsx veya .csv formatında yükleyin.\n"+e.message);
    }finally{
      setLoading(false);
    }
  };

  const buildPreview=()=>{
    const poolByBarcode=new Map(pool.map(p=>[p.barcode,p]));
    const prev=rows.slice(0,8).map(row=>{
      const barcode=String(row[parseInt(map.barcode)]??"").trim();
      const name=String(row[parseInt(map.name)]??"").trim();
      const price=parseFloat(String(row[parseInt(map.price)]??"").replace(",","."))||0;
      const brand=map.brand!==""?String(row[parseInt(map.brand)]??"").trim():"";
      return{barcode,name,price,brand,exists:barcode&&poolByBarcode.has(barcode)};
    }).filter(r=>r.name||r.barcode);
    setPreview(prev);
    setStep("preview");
  };

  const doImport=async()=>{
    setImporting(true);
    const poolByBarcode=new Map(pool.map(p=>[p.barcode,p]));
    let added=0,updated=0,skipped=0;
    const priceUpdates=[]; // {id, basePrice}
    const newRows=[]; // insert edilecek satırlar (DB formatı)

    rows.forEach(row=>{
      const barcode=String(row[parseInt(map.barcode)]??"").trim();
      const name=String(row[parseInt(map.name)]??"").trim();
      const price=parseFloat(String(row[parseInt(map.price)]??"").replace(",","."))||0;
      const brand=map.brand!==""?String(row[parseInt(map.brand)]??"").trim():"";

      if(!name&&!barcode){skipped++;return;}

      const existing=barcode?poolByBarcode.get(barcode):null;
      if(existing){
        if(price>0){
          priceUpdates.push({id:existing.id,basePrice:price});
          updated++;
        } else {
          skipped++;
        }
      } else {
        newRows.push({
          barcode:barcode||null, brand:brand||"Bilinmiyor", name:name||"İsimsiz Ürün",
          category:"nemlendirici", usage:"both", base_price:price,
          actives:[], pairs_with:[], certs:[], skin_types:[], concerns:[],
          boycott:"bilinmiyor", pregnancy_safe:"unknown",
          description:"Excel'den içe aktarıldı.", how_to:"",
        });
        added++;
      }
    });

    // Fiyat güncellemelerini sırayla uygula
    for(const u of priceUpdates){
      await supabase.from("pool").update({base_price:u.basePrice}).eq("id",u.id);
    }
    // Yeni ürünleri toplu ekle
    let insertedRows=[];
    if(newRows.length>0){
      const{data,error}=await supabase.from("pool").insert(newRows).select();
      if(!error&&data)insertedRows=data;
    }

    const merged=pool.map(p=>{
      const u=priceUpdates.find(x=>x.id===p.id);
      return u?{...p,basePrice:u.basePrice}:p;
    });
    onImport([...merged,...insertedRows.map(poolRowToJs)]);
    setResult({added,updated,skipped});
    setImporting(false);
    setStep("done");
  };

  const reset=()=>{
    setRows([]);setHeaders([]);setPreview([]);
    setMap({barcode:"",name:"",price:"",brand:""});
    setStep("upload");setResult({added:0,updated:0,skipped:0});
    if(fileRef.current)fileRef.current.value="";
  };

  const FIELD_LABELS={barcode:"Barkod",name:"Ürün Adı",price:"Fiyat",brand:"Marka (opsiyonel)"};

  return(
    <div style={{padding:"0 0 20px"}}>
      {step==="upload"&&(
        <div>
          <div style={{background:"#F8F9FA",borderRadius:12,border:"2px dashed #CCC",
            padding:24,textAlign:"center",marginBottom:12}}>
            <div style={{fontSize:32,marginBottom:8}}>📊</div>
            <div style={{fontSize:14,fontWeight:600,color:"#333",marginBottom:4}}>Excel veya CSV yükle</div>
            <div style={{fontSize:12,color:"#888",marginBottom:14,lineHeight:1.5}}>
              Barkod, ürün adı ve fiyat sütunları otomatik tanınır.<br/>
              Eşleşen barkodların fiyatı güncellenir, yeni barkodlar havuza eklenir.<br/>
              .xlsx · .xls · .csv formatları desteklenir.
            </div>
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv"
              onChange={e=>handleFile(e.target.files[0])} style={{display:"none"}}/>
            <button type="button" onClick={()=>fileRef.current&&fileRef.current.click()}
              style={{display:"inline-block",background:"#0D0D0D",color:"#fff",border:"none",
                borderRadius:10,padding:"10px 24px",fontSize:13,fontWeight:600,cursor:"pointer"}}>
              {loading?"⏳ Yükleniyor…":"Dosya Seç"}
            </button>
          </div>
          <div style={{background:"#EEF0FB",borderRadius:10,padding:12}}>
            <div style={{fontSize:11,fontWeight:600,color:"#3949AB",marginBottom:6}}>
              📋 Beklenen sütun formatı (örnek):
            </div>
            <div style={{fontSize:11,fontFamily:"monospace",color:"#555",lineHeight:1.8}}>
              | Barkod &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;| Ürün Adı &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;| Fiyat |<br/>
              | 3282770203745 | Avène Temizleyici | 520 &nbsp;&nbsp;|<br/>
              | 3522930002285 | Caudalie Serum &nbsp;&nbsp;&nbsp;| 1250 &nbsp;|
            </div>
          </div>
        </div>
      )}

      {step==="map"&&(
        <div>
          <div style={{fontSize:13,fontWeight:600,color:"#1C1C1A",marginBottom:4}}>Sütun Eşleştirme</div>
          <div style={{fontSize:11,color:"#888",marginBottom:14}}>
            {rows.length} satır okundu. Hangi sütun hangi alana karşılık geliyor?
          </div>
          {Object.entries(FIELD_LABELS).map(([field,label])=>(
            <div key={field} style={{marginBottom:10}}>
              <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:4}}>
                {label}{field!=="brand"&&<span style={{color:"#C0392B"}}> *</span>}
              </div>
              <select value={map[field]} onChange={e=>setMap(p=>({...p,[field]:e.target.value}))}
                style={{width:"100%",padding:"8px 10px",border:"1px solid #E0E0E0",
                  borderRadius:8,fontSize:13,color:"#333",background:"#fff"}}>
                <option value="">— Seç —</option>
                {headers.map((h,i)=>(
                  <option key={i} value={String(i)}>{h} (örn: {String(rows[0]?.[i]??"")})</option>
                ))}
              </select>
            </div>
          ))}
          <div style={{display:"flex",gap:8,marginTop:16}}>
            <button onClick={reset}
              style={{flex:1,background:"#F0F0F0",color:"#333",border:"none",
                borderRadius:10,padding:"10px",fontSize:13,cursor:"pointer"}}>← Geri</button>
            <button onClick={buildPreview} disabled={!map.name&&!map.barcode}
              style={{flex:2,background:map.name||map.barcode?"#0D0D0D":"#CCC",color:"#fff",
                border:"none",borderRadius:10,padding:"10px",fontSize:13,fontWeight:600,
                cursor:map.name||map.barcode?"pointer":"default"}}>Önizle →</button>
          </div>
        </div>
      )}

      {step==="preview"&&(
        <div>
          <div style={{fontSize:13,fontWeight:600,color:"#1C1C1A",marginBottom:4}}>Önizleme — İlk 8 Satır</div>
          <div style={{fontSize:11,color:"#888",marginBottom:10}}>Toplam {rows.length} satır işlenecek.</div>
          <div style={{overflowX:"auto",marginBottom:14}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
              <thead>
                <tr style={{background:"#F5F5F5"}}>
                  {["Barkod","Ürün Adı","Fiyat (₺)","Durum"].map(h=>(
                    <th key={h} style={{padding:"6px 8px",textAlign:"left",fontWeight:600,color:"#555",
                      whiteSpace:"nowrap",borderBottom:"1px solid #E0E0E0"}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {preview.map((r,i)=>(
                  <tr key={i} style={{borderBottom:"1px solid #F0F0F0"}}>
                    <td style={{padding:"5px 8px",color:"#888",fontFamily:"monospace",fontSize:10}}>{r.barcode||"—"}</td>
                    <td style={{padding:"5px 8px",color:"#1C1C1A",maxWidth:150,overflow:"hidden",
                      textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.name||"—"}</td>
                    <td style={{padding:"5px 8px",fontWeight:600,color:"#2C4A3E"}}>
                      {r.price>0?r.price.toLocaleString("tr-TR")+" ₺":"—"}
                    </td>
                    <td style={{padding:"5px 8px",color:r.exists?"#7D4700":"#1A237E",fontWeight:600,whiteSpace:"nowrap"}}>
                      {r.exists?"🔄 Güncelle":"➕ Yeni"}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{background:"#EEF0FB",borderRadius:8,padding:10,marginBottom:14,fontSize:11,color:"#333",lineHeight:1.6}}>
            🔄 <b>Güncelle</b>: Barkod havuzda var, fiyatı güncellenir.<br/>
            ➕ <b>Yeni</b>: Minimum bilgiyle havuza eklenir, sonra düzenleyebilirsiniz (cilt tipi/endişe etiketleri boş gelir).
          </div>
          <div style={{display:"flex",gap:8}}>
            <button onClick={()=>setStep("map")} disabled={importing}
              style={{flex:1,background:"#F0F0F0",color:"#333",border:"none",
                borderRadius:10,padding:"10px",fontSize:13,cursor:importing?"default":"pointer"}}>← Geri</button>
            <button onClick={doImport} disabled={importing}
              style={{flex:2,background:importing?"#999":"#2C4A3E",color:"#fff",border:"none",
                borderRadius:10,padding:"10px",fontSize:13,fontWeight:700,cursor:importing?"default":"pointer"}}>
              {importing?"Aktarılıyor…":`${rows.length} Ürünü Havuza Aktar →`}
            </button>
          </div>
          {importing&&<div style={{fontSize:10,color:"#999",textAlign:"center",marginTop:6}}>
            Bu işlem satır sayısına göre biraz sürebilir, lütfen sayfadan ayrılmayın.
          </div>}
        </div>
      )}

      {step==="done"&&(
        <div style={{textAlign:"center"}}>
          <div style={{fontSize:40,marginBottom:12}}>🎉</div>
          <div style={{fontSize:16,fontWeight:700,color:"#1C1C1A",marginBottom:16}}>İçe aktarma tamamlandı</div>
          <div style={{display:"flex",gap:8,marginBottom:20}}>
            {[
              {label:"Eklendi",value:result.added,bg:"#E9F7EF",c:"#1E6B3E"},
              {label:"Güncellendi",value:result.updated,bg:"#FEF3E0",c:"#7D4700"},
              {label:"Atlandı",value:result.skipped,bg:"#F4F4F4",c:"#666"},
            ].map(({label,value,bg,c})=>(
              <div key={label} style={{flex:1,background:bg,borderRadius:10,padding:"12px 8px",textAlign:"center"}}>
                <div style={{fontSize:24,fontWeight:700,color:c}}>{value}</div>
                <div style={{fontSize:10,color:c,marginTop:2}}>{label}</div>
              </div>
            ))}
          </div>
          <button onClick={reset}
            style={{width:"100%",background:"#0D0D0D",color:"#fff",border:"none",
              borderRadius:10,padding:"11px",fontSize:13,fontWeight:600,cursor:"pointer"}}>
            Yeni Dosya Yükle
          </button>
        </div>
      )}
    </div>
  );
}

function SuperAdmin({pool,setPool,accounts,onApprove,onReject,onSuspend,onReactivate,onDeleteAccount,onResetPassword,superadminAccount,onUpdateSuperadmin,onBack}){
  // Güvenlik: 10 dakika etkileşim olmazsa panelden otomatik çıkış yapılır
  useIdleTimer(10*60*1000,onBack,true);
  const [panelTab,setPanelTab]=useState("pool"); // pool | excel | accounts
  const [search,setSearch]=useState("");
  const [filterBrand,setFilterBrand]=useState("");
  const [filterCat,setFilterCat]=useState("");
  const [sortBy,setSortBy]=useState("name");
  const [selected,setSelected]=useState(new Set());
  const [msg,setMsg]=useState("");
  const [showAdd,setShowAdd]=useState(false);
  const [showCredModal,setShowCredModal]=useState(false);
  const [showFillScanner,setShowFillScanner]=useState(false);
  const pendingCount=accounts.filter(a=>a.status==="pending").length;
  const toggleSelect=id=>setSelected(prev=>{const n=new Set(prev);n.has(id)?n.delete(id):n.add(id);return n;});
  const [form,setForm]=useState({brand:"",name:"",barcode:"",photo:"",category:"temizleyici",usage:"both",
    basePrice:"",actives:"",desc:"",how_to:"",boycott:"bilinmiyor",pregnancy_safe:"unknown",
    skin_types:[],concerns:[]});
  const toggleFormST=(st)=>setForm(p=>({...p,skin_types:p.skin_types.includes(st)?p.skin_types.filter(x=>x!==st):[...p.skin_types,st]}));
  const toggleFormC=(c)=>setForm(p=>({...p,concerns:p.concerns.includes(c)?p.concerns.filter(x=>x!==c):[...p.concerns,c]}));
  const [errors,setErrors]=useState({});

  const notify=m=>{setMsg(m);setTimeout(()=>setMsg(""),2500);};
  const updatePool=async(id,f,v)=>{
    setPool(prev=>prev.map(p=>p.id===id?{...p,[f]:v}:p)); // iyimser güncelleme
    const dbField=POOL_FIELD_MAP[f]||f;
    const{error}=await supabase.from("pool").update({[dbField]:v}).eq("id",id);
    if(error)notify("⚠ Kaydedilemedi: "+error.message);
  };
  const removeFromPool=async id=>{
    setPool(prev=>prev.filter(p=>p.id!==id));
    const{error}=await supabase.from("pool").delete().eq("id",id);
    notify(error?"⚠ Silinemedi: "+error.message:"Ürün silindi.");
  };

  const validate=()=>{
    const e={};
    if(!form.brand.trim())e.brand="Marka gerekli";
    if(!form.name.trim())e.name="Ürün adı gerekli";
    if(!form.basePrice||parseInt(form.basePrice)<=0)e.basePrice="Geçerli fiyat giriniz";
    if(!form.desc.trim())e.desc="Açıklama gerekli";
    setErrors(e);return Object.keys(e).length===0;
  };

  const addToPool=async()=>{
    if(!validate())return;
    const row={
      barcode:form.barcode.trim()||null, brand:form.brand.trim(), name:form.name.trim(),
      category:form.category, usage:form.usage, base_price:parseInt(form.basePrice),
      actives:form.actives.split(",").map(s=>s.trim()).filter(Boolean),
      pairs_with:[], certs:[], skin_types:form.skin_types, concerns:form.concerns,
      boycott:form.boycott, pregnancy_safe:form.pregnancy_safe,
      photo:form.photo.trim()||null, description:form.desc.trim(), how_to:form.how_to.trim(),
    };
    const{data,error}=await supabase.from("pool").insert(row).select().single();
    if(error){notify("⚠ Havuza eklenemedi: "+error.message);return;}
    setPool(prev=>[...prev,poolRowToJs(data)]);
    setForm({brand:"",name:"",barcode:"",photo:"",category:"temizleyici",usage:"both",
      basePrice:"",actives:"",desc:"",how_to:"",boycott:"bilinmiyor",pregnancy_safe:"unknown",
      skin_types:[],concerns:[]});
    setErrors({});setShowAdd(false);notify("✓ Ürün havuza eklendi.");
  };

  const F=({f,l,ph,type="text"})=>(
    <div style={{marginBottom:8}}>
      <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:3}}>{l}</div>
      <input value={form[f]} onChange={e=>setForm(prev=>({...prev,[f]:e.target.value}))} placeholder={ph} type={type}
        style={{width:"100%",padding:"7px 10px",border:`1px solid ${errors[f]?"#C0392B":"#E0E0E0"}`,
          borderRadius:7,fontSize:12,outline:"none",boxSizing:"border-box"}}/>
      {errors[f]&&<div style={{color:"#C0392B",fontSize:10,marginTop:2}}>⚠ {errors[f]}</div>}
    </div>
  );

  const poolBrands=[...new Set(pool.map(p=>p.brand))].sort((a,b)=>a.localeCompare(b,"tr"));
  const poolSortFn=(a,b)=>{
    if(sortBy==="price_desc")return b.basePrice-a.basePrice;
    if(sortBy==="price_asc")return a.basePrice-b.basePrice;
    return a.name.localeCompare(b.name,"tr");
  };
  const filtered=pool
    .filter(p=>filterBrand===""||p.brand===filterBrand)
    .filter(p=>filterCat===""||p.category===filterCat)
    .filter(p=>search===""||
      p.name.toLowerCase().includes(search.toLowerCase())||p.brand.toLowerCase().includes(search.toLowerCase()))
    .sort(poolSortFn);

  const bulkDeletePool=async()=>{
    if(selected.size===0)return;
    if(!window.confirm(`${selected.size} ürün havuzdan kalıcı olarak silinsin mi? Bu ürünleri kataloğuna eklemiş eczaneler etkilenebilir.`))return;
    const ids=[...selected];
    setPool(prev=>prev.filter(p=>!ids.includes(p.id)));
    await supabase.from("pool").delete().in("id",ids);
    notify(`${ids.length} ürün silindi.`);
    setSelected(new Set());
  };

  const bulkAnalyzePool=async()=>{
    if(selected.size===0)return;
    const targets=pool.filter(p=>selected.has(p.id));
    let applied=0,skipped=0;
    for(const item of targets){
      const result=analyzeIngredients(item.actives);
      if(result.matched.length===0){skipped++;continue;}
      setPool(prev=>prev.map(p=>p.id===item.id?{...p,skin_types:result.skin_types,concerns:result.concerns,
        pregnancy_safe:result.pregnancy_safe!=="unknown"?result.pregnancy_safe:p.pregnancy_safe}:p));
      const dbUpdate={skin_types:result.skin_types,concerns:result.concerns};
      if(result.pregnancy_safe!=="unknown")dbUpdate.pregnancy_safe=result.pregnancy_safe;
      await supabase.from("pool").update(dbUpdate).eq("id",item.id);
      applied++;
    }
    notify(`🧪 ${applied} ürün analiz edildi${skipped>0?`, ${skipped} ürün için tanınan içerik bulunamadı`:""}.`);
    setSelected(new Set());
  };

  return(
    <div className="admin-shell" style={{background:"#F5F5F5",minHeight:"100vh",
      paddingBottom:20,fontFamily:"'Inter',sans-serif"}}>
      {showCredModal&&<ChangeCredentialsModal title="Süper Admin Hesabı"
        initialEmail={superadminAccount.email} showPhone={false}
        onSave={fields=>onUpdateSuperadmin(fields)} onClose={()=>setShowCredModal(false)}/>}

      {showFillScanner&&<BarcodeScanner
        mode="fill"
        onFill={code=>setForm(p=>({...p,barcode:code}))}
        onClose={()=>setShowFillScanner(false)}/>}

      <div style={{background:"#0D0D0D",color:"#fff",padding:"14px 16px 0",position:"sticky",top:0,zIndex:50}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
          <div>
            <div style={{fontSize:16,fontWeight:700}}>🔐 Süper Admin</div>
            <div style={{fontSize:11,opacity:.5,marginTop:1}}>{pool.length} ürün · {[...new Set(pool.map(p=>p.brand))].length} marka</div>
          </div>
          <div style={{display:"flex",gap:7}}>
            <button onClick={()=>setShowCredModal(true)}
              style={{background:"rgba(255,255,255,.08)",border:"none",color:"rgba(255,255,255,.6)",
                borderRadius:8,padding:"6px 10px",fontSize:11,cursor:"pointer"}}>👤 Hesabım</button>
            <button onClick={onBack} style={{background:"rgba(255,255,255,.1)",border:"none",color:"#fff",
              borderRadius:8,padding:"6px 10px",fontSize:12,cursor:"pointer"}}>Çıkış</button>
          </div>
        </div>
        <div style={{display:"flex",gap:4}}>
          {[{id:"pool",l:"Master Havuz"},{id:"excel",l:"Excel İçe Aktar"},{id:"accounts",l:`Hesap Talepleri${pendingCount>0?` (${pendingCount})`:""}`}].map(t=>(
            <button key={t.id} onClick={()=>setPanelTab(t.id)}
              style={{flex:1,background:panelTab===t.id?"rgba(255,255,255,.15)":"transparent",border:"none",
                color:panelTab===t.id?"#fff":"rgba(255,255,255,.5)",borderRadius:"8px 8px 0 0",
                padding:"7px 4px",fontSize:11,fontWeight:panelTab===t.id?700:400,cursor:"pointer"}}>
              {t.l}{t.id==="accounts"&&pendingCount>0&&panelTab!==t.id&&
                <span style={{marginLeft:4,background:"#C0392B",color:"#fff",borderRadius:8,
                  padding:"1px 6px",fontSize:9,fontWeight:700}}>{pendingCount}</span>}
            </button>
          ))}
        </div>
      </div>
      <div style={{padding:"12px 14px 0"}}>
        <Toast msg={msg}/>
        {panelTab==="accounts"&&(
          <AccountsPanel accounts={accounts} onApprove={onApprove} onReject={onReject}
            onSuspend={onSuspend} onReactivate={onReactivate} onDelete={onDeleteAccount}
            onResetPassword={onResetPassword} notify={notify}/>
        )}
        {panelTab==="excel"&&(
          <PoolExcelImport pool={pool} onImport={newPool=>{setPool(newPool);notify("✓ Havuz Excel'den güncellendi.");}}/>
        )}
        {panelTab==="pool"&&(<>
        <div style={{display:"flex",gap:8,marginBottom:10}}>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Ara…"
            style={{flex:1,padding:"9px 12px",border:"1.5px solid #E0E0E0",borderRadius:10,fontSize:13,outline:"none"}}/>
          <button onClick={()=>{setShowAdd(!showAdd);setErrors({});}}
            style={{background:"#0D0D0D",color:"#fff",border:"none",borderRadius:10,
              padding:"9px 16px",fontSize:13,fontWeight:600,cursor:"pointer"}}>
            {showAdd?"✕":"+ Ekle"}
          </button>
        </div>
        <FilterChips label="Marka" options={poolBrands} value={filterBrand} onChange={setFilterBrand}/>
        <FilterChips label="Kategori" options={Object.keys(CAT_LABELS)} value={filterCat}
          onChange={setFilterCat} getLabel={c=>CAT_LABELS[c]}/>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
          <div style={{fontSize:11,color:"#999"}}>{filtered.length} / {pool.length} ürün · {poolBrands.length} marka</div>
          <select value={sortBy} onChange={e=>setSortBy(e.target.value)}
            style={{padding:"6px 8px",border:"1px solid #E0E0E0",borderRadius:8,fontSize:11,
              background:"#fff",color:"#555"}}>
            <option value="name">Ada göre (A-Z)</option>
            <option value="price_asc">Fiyata göre (artan)</option>
            <option value="price_desc">Fiyata göre (azalan)</option>
          </select>
        </div>
        {selected.size>0&&(
          <div style={{background:"#0D0D0D",borderRadius:12,padding:"10px 12px",marginBottom:10,
            display:"flex",alignItems:"center",gap:8}}>
            <span style={{color:"#fff",fontSize:12,fontWeight:600}}>{selected.size} seçili</span>
            <button onClick={bulkAnalyzePool}
              style={{background:"#EEF0FB",color:"#3949AB",border:"none",borderRadius:7,
                padding:"6px 10px",fontSize:11,fontWeight:600,cursor:"pointer"}}>🧪 Analiz Et</button>
            <button onClick={bulkDeletePool}
              style={{background:"#FDECEA",color:"#C0392B",border:"none",borderRadius:7,
                padding:"6px 10px",fontSize:11,fontWeight:600,cursor:"pointer"}}>🗑 Sil</button>
            <button onClick={()=>setSelected(new Set())}
              style={{background:"none",color:"#fff",border:"none",fontSize:11,cursor:"pointer",
                marginLeft:"auto",opacity:.7}}>Seçimi Temizle</button>
          </div>
        )}
        {showAdd&&(
          <div style={{background:"#fff",borderRadius:14,border:"1px solid #E0E0E0",padding:14,marginBottom:14}}>
            <div style={{fontSize:14,fontWeight:700,color:"#1C1C1A",marginBottom:4}}>Yeni Ürün Ekle</div>
            <div style={{fontSize:11,color:"#999",marginBottom:12}}>
              Çok sayıda ürün ekleyecekseniz "Excel İçe Aktar" sekmesi çok daha hızlıdır.
            </div>
            <F f="brand" l="Marka *" ph="Avène"/>
            <F f="name" l="Ürün Adı *" ph="Temizleyici Jel"/>
            <F f="barcode" l="Barkod" ph="8690123456789"/>
            <F f="photo" l="Fotoğraf URL (opsiyonel)" ph="https://..."/>
            {/* Barkod — scan butonu ile */}
            <div style={{marginBottom:8}}>
              <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:3}}>Barkod (EAN-13)</div>
              <div style={{display:"flex",gap:6}}>
                <input value={form.barcode} onChange={e=>setForm(p=>({...p,barcode:e.target.value}))}
                  placeholder="3282770203745" inputMode="numeric"
                  style={{flex:1,padding:"7px 10px",border:"1px solid #E0E0E0",borderRadius:7,
                    fontSize:12,outline:"none",boxSizing:"border-box"}}/>
                <button type="button" onClick={()=>setShowFillScanner(true)}
                  style={{background:"#1C1C1A",color:"#fff",border:"none",borderRadius:7,
                    padding:"7px 12px",fontSize:13,cursor:"pointer",flexShrink:0,fontWeight:600}}
                  title="Barkod tara">
                  📷
                </button>
              </div>
              {form.barcode&&<div style={{fontSize:10,color:"#1E6B3E",marginTop:3}}>
                ✓ Barkod girildi: {form.barcode}
              </div>}
            </div>
            <F f="basePrice" l="Taban Fiyat (₺) *" ph="520" type="number"/>
            <F f="actives" l="Aktif İçerikler (virgülle)" ph="Hyalüronik Asit, Niasinamid"/>
            <IngredientAnalysisBox actives={form.actives} onApply={result=>setForm(p=>({
              ...p, skin_types:result.skin_types, concerns:result.concerns,
              pregnancy_safe:result.pregnancy_safe!=="unknown"?result.pregnancy_safe:p.pregnancy_safe,
            }))}/>
            <div style={{height:10}}/>
            <F f="desc" l="Açıklama *" ph="Kısa ürün açıklaması"/>
            <F f="how_to" l="Nasıl Kullanılır" ph="Yüze uygulayın…"/>
            <div style={{display:"flex",gap:8,marginBottom:8}}>
              <div style={{flex:1}}>
                <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:3}}>Kategori</div>
                <select value={form.category} onChange={e=>setForm(p=>({...p,category:e.target.value}))}
                  style={{width:"100%",padding:"7px 10px",border:"1px solid #E0E0E0",borderRadius:7,fontSize:12}}>
                  {Object.entries(CAT_LABELS).map(([k,v])=><option key={k} value={k}>{v}</option>)}
                </select>
              </div>
              <div style={{flex:1}}>
                <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:3}}>Kullanım</div>
                <select value={form.usage} onChange={e=>setForm(p=>({...p,usage:e.target.value}))}
                  style={{width:"100%",padding:"7px 10px",border:"1px solid #E0E0E0",borderRadius:7,fontSize:12}}>
                  <option value="both">Sabah & Akşam</option><option value="am">Sabah</option><option value="pm">Akşam</option>
                </select>
              </div>
            </div>
            <div style={{display:"flex",gap:8,marginBottom:10}}>
              <div style={{flex:1}}>
                <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:3}}>Boykot</div>
                <select value={form.boycott} onChange={e=>setForm(p=>({...p,boycott:e.target.value}))}
                  style={{width:"100%",padding:"7px 10px",border:"1px solid #E0E0E0",borderRadius:7,fontSize:12}}>
                  <option value="temiz">✓ Boykot Yok</option><option value="boykot">⛔ Boykot</option><option value="bilinmiyor">? Bilinmiyor</option>
                </select>
              </div>
              <div style={{flex:1}}>
                <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:3}}>Hamile Güvenliği</div>
                <select value={form.pregnancy_safe} onChange={e=>setForm(p=>({...p,pregnancy_safe:e.target.value}))}
                  style={{width:"100%",padding:"7px 10px",border:"1px solid #E0E0E0",borderRadius:7,fontSize:12}}>
                  <option value="safe">✓ Güvenli</option><option value="avoid">⛔ Kaçınılmalı</option>
                  <option value="consult">⚠ Doktora Danış</option><option value="unknown">? Bilinmiyor</option>
                </select>
              </div>
            </div>
            {/* Cilt tipi uyumu */}
            <div style={{marginBottom:10}}>
              <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:6}}>
                Cilt Tipi Uyumu <span style={{color:"#C0392B"}}>*</span>
                <span style={{fontWeight:400,color:"#999",marginLeft:4}}>(algoritma için seç)</span>
              </div>
              <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                {Object.entries(SKIN_TYPE_LABELS).map(([k,v])=>{
                  const has=form.skin_types.includes(k);
                  return <button key={k} type="button" onClick={()=>toggleFormST(k)}
                    style={{padding:"5px 12px",borderRadius:8,fontSize:12,fontWeight:500,cursor:"pointer",
                      border:`1.5px solid ${has?"#2C4A3E":"#E0E0E0"}`,
                      background:has?"#E9F7EF":"#FAFAFA",color:has?"#2C4A3E":"#888"}}>
                    {has?"✓ ":""}{v}
                  </button>;
                })}
              </div>
              {form.skin_types.length===0&&<div style={{fontSize:10,color:"#E67E22",marginTop:4}}>
                ⚠ En az bir cilt tipi seç — algoritma bu ürünü gösteremez
              </div>}
            </div>

            {/* Endişe etiketleri */}
            <div style={{marginBottom:10}}>
              <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:6}}>
                Endişe Etiketleri <span style={{color:"#C0392B"}}>*</span>
                <span style={{fontWeight:400,color:"#999",marginLeft:4}}>(algoritma için seç)</span>
              </div>
              <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                {Object.entries(CONCERN_LABELS).map(([k,v])=>{
                  const has=form.concerns.includes(k);
                  return <button key={k} type="button" onClick={()=>toggleFormC(k)}
                    style={{padding:"5px 12px",borderRadius:8,fontSize:12,fontWeight:500,cursor:"pointer",
                      border:`1.5px solid ${has?"#C4974A":"#E0E0E0"}`,
                      background:has?"#FEF3E0":"#FAFAFA",color:has?"#7D4700":"#888"}}>
                    {has?"✓ ":""}{v}
                  </button>;
                })}
              </div>
              {form.concerns.length===0&&<div style={{fontSize:10,color:"#E67E22",marginTop:4}}>
                ⚠ En az bir endişe seç — müşteri eşleşmesi yapılamaz
              </div>}
            </div>

            <button onClick={addToPool}
              style={{width:"100%",background:"#0D0D0D",color:"#fff",border:"none",borderRadius:9,
                padding:11,fontSize:13,fontWeight:700,cursor:"pointer"}}>Havuza Ekle →</button>
          </div>
        )}
        <div className="admin-grid">
          {filtered.map(p=>(
            <PoolItem key={p.id} product={p} onUpdate={updatePool} onRemove={removeFromPool}
              selected={selected.has(p.id)} onToggleSelect={()=>toggleSelect(p.id)}/>
          ))}
        </div>
        </>)}
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════
// HESAP TALEPLERİ / HESAP YÖNETİMİ (Süper Admin)
// ══════════════════════════════════════════════════════════════
const ACCOUNT_STATUS_META = {
  pending:  {label:"⏳ Onay Bekliyor", bg:"#FFF3E0",color:"#7D4700"},
  approved: {label:"✓ Onaylı",         bg:"#E9F7EF",color:"#1E6B3E"},
  rejected: {label:"⛔ Reddedildi",     bg:"#FDECEA",color:"#8B2E2E"},
  suspended:{label:"⏸ Askıda",         bg:"#F4F4F4",color:"#666"},
};

function AccountCard({acc,onApprove,onReject,onSuspend,onReactivate,onDelete,onResetPassword,notify}){
  const meta=ACCOUNT_STATUS_META[acc.status]||ACCOUNT_STATUS_META.pending;
  const created=acc.createdAt?new Date(acc.createdAt).toLocaleString("tr-TR"):"";

  const doReset=async()=>{
    if(!window.confirm(`${acc.pharmacyName} (${acc.email}) adresine şifre sıfırlama bağlantısı gönderilsin mi?`))return;
    await onResetPassword(acc.id);
    notify(`✉️ Şifre sıfırlama bağlantısı ${acc.email} adresine gönderildi.`);
  };
  const doDelete=async()=>{
    if(!window.confirm(`${acc.pharmacyName} hesabı kalıcı olarak silinsin mi?`))return;
    await onDelete(acc.id);
  };

  return(
    <div style={{background:"#fff",borderRadius:12,border:"1px solid #E0E0E0",padding:12}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8,marginBottom:6}}>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:13,fontWeight:700,color:"#1C1C1A"}}>{acc.pharmacyName}</div>
          <div style={{fontSize:11,color:"#888",marginTop:1}}>{acc.contactName}</div>
        </div>
        <span style={{background:meta.bg,color:meta.color,fontSize:10,fontWeight:700,
          padding:"3px 8px",borderRadius:6,whiteSpace:"nowrap"}}>{meta.label}</span>
      </div>
      <div style={{fontSize:11,color:"#555",lineHeight:1.7}}>
        <div>📧 {acc.email}</div>
        <div>📞 {acc.phone}</div>
        {created&&<div style={{color:"#999"}}>Talep: {created}</div>}
      </div>
      <div style={{display:"flex",flexWrap:"wrap",gap:6,marginTop:10}}>
        {acc.status==="pending"&&<>
          <button onClick={()=>onApprove(acc.id)}
            style={{flex:1,background:"#1E6B3E",color:"#fff",border:"none",borderRadius:8,
              padding:"8px 0",fontSize:12,fontWeight:600,cursor:"pointer"}}>✓ Onayla</button>
          <button onClick={()=>onReject(acc.id)}
            style={{flex:1,background:"#FDECEA",color:"#C0392B",border:"none",borderRadius:8,
              padding:"8px 0",fontSize:12,fontWeight:600,cursor:"pointer"}}>⛔ Reddet</button>
        </>}
        {acc.status==="approved"&&<>
          <button onClick={doReset}
            style={{flex:1,background:"#F0F0F0",color:"#333",border:"none",borderRadius:8,
              padding:"8px 0",fontSize:12,fontWeight:600,cursor:"pointer"}}>✉️ Şifre Sıfırlama Maili</button>
          <button onClick={()=>onSuspend(acc.id)}
            style={{flex:1,background:"#FFF3E0",color:"#7D4700",border:"none",borderRadius:8,
              padding:"8px 0",fontSize:12,fontWeight:600,cursor:"pointer"}}>⏸ Askıya Al</button>
        </>}
        {acc.status==="rejected"&&<button onClick={()=>onApprove(acc.id)}
          style={{flex:1,background:"#1E6B3E",color:"#fff",border:"none",borderRadius:8,
            padding:"8px 0",fontSize:12,fontWeight:600,cursor:"pointer"}}>✓ Yine de Onayla</button>}
        {acc.status==="suspended"&&<button onClick={()=>onReactivate(acc.id)}
          style={{flex:1,background:"#1E6B3E",color:"#fff",border:"none",borderRadius:8,
            padding:"8px 0",fontSize:12,fontWeight:600,cursor:"pointer"}}>▶ Aktifleştir</button>}
        <button onClick={doDelete}
          style={{background:"#FDECEA",color:"#C0392B",border:"none",borderRadius:8,
            padding:"8px 12px",fontSize:12,fontWeight:600,cursor:"pointer"}}>🗑</button>
      </div>
    </div>
  );
}

function AccountsPanel({accounts,onApprove,onReject,onSuspend,onReactivate,onDelete,onResetPassword,notify}){
  const [filter,setFilter]=useState("all");
  const order={pending:0,approved:1,suspended:2,rejected:3};
  const list=[...accounts]
    .filter(a=>filter==="all"||a.status===filter)
    .sort((a,b)=>(order[a.status]??9)-(order[b.status]??9));

  return(
    <div>
      <div style={{fontSize:11,color:"#999",marginBottom:8,lineHeight:1.5}}>
        Yeni eczane kayıt talepleri burada listelenir. Onaylanmayan hesaplar eczane paneline giriş yapamaz.
      </div>
      <div style={{display:"flex",gap:6,marginBottom:10,flexWrap:"wrap"}}>
        {[{id:"all",l:"Tümü"},{id:"pending",l:"Bekleyen"},{id:"approved",l:"Onaylı"},
          {id:"suspended",l:"Askıda"},{id:"rejected",l:"Reddedilen"}].map(f=>(
          <button key={f.id} onClick={()=>setFilter(f.id)}
            style={{background:filter===f.id?"#0D0D0D":"#fff",color:filter===f.id?"#fff":"#555",
              border:"1px solid #E0E0E0",borderRadius:8,padding:"5px 11px",fontSize:11,
              fontWeight:filter===f.id?700:500,cursor:"pointer"}}>{f.l}</button>
        ))}
      </div>
      {list.length===0
        ?<div style={{textAlign:"center",padding:40,color:"#999",fontSize:13}}>Bu filtrede hesap yok.</div>
        :<div style={{display:"flex",flexDirection:"column",gap:8,paddingBottom:20}}>
          {list.map(acc=>(
            <AccountCard key={acc.id} acc={acc} onApprove={onApprove} onReject={onReject}
              onSuspend={onSuspend} onReactivate={onReactivate} onDelete={onDelete}
              onResetPassword={onResetPassword} notify={notify}/>
          ))}
        </div>}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════
// ADMİN ROL SEÇİM
// ══════════════════════════════════════════════════════════════
function AdminSelect({onSelect,onBack}){
  return(
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",
      background:"#0D0D0D",fontFamily:"'Inter',sans-serif"}}>
      <div style={{maxWidth:360,width:"90%",textAlign:"center"}}>
        <div style={{fontSize:13,color:"rgba(255,255,255,.3)",marginBottom:6,letterSpacing:2}}>YÖNETİM PANELİ</div>
        <div style={{fontSize:24,fontWeight:700,color:"#fff",marginBottom:32}}>Giriş türünü seçin</div>
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {[
            {role:"pharmacy",icon:"🏪",title:"Eczane Yöneticisi",desc:"Katalog, fiyat, indirim, stok",bg:"rgba(255,255,255,.08)"},
            {role:"superadmin",icon:"🔐",title:"Süper Admin",desc:"Master ürün havuzu",bg:"rgba(255,255,255,.05)"},
          ].map(({role,icon,title,desc,bg})=>(
            <button key={role} onClick={()=>onSelect(role)}
              style={{background:bg,border:"1px solid rgba(255,255,255,.1)",color:"#fff",borderRadius:14,
                padding:"18px 20px",textAlign:"left",cursor:"pointer",display:"flex",alignItems:"center",gap:14}}>
              <span style={{fontSize:28}}>{icon}</span>
              <div>
                <div style={{fontSize:15,fontWeight:600}}>{title}</div>
                <div style={{fontSize:12,opacity:.5,marginTop:2}}>{desc}</div>
              </div>
              <span style={{marginLeft:"auto",opacity:.3,fontSize:18}}>→</span>
            </button>
          ))}
        </div>
        <button onClick={onBack} style={{marginTop:20,background:"none",border:"none",
          color:"rgba(255,255,255,.3)",fontSize:12,cursor:"pointer"}}>Müşteri ekranına dön</button>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════
// ANA UYGULAMA
// ══════════════════════════════════════════════════════════════
export default function App(){
  useResponsiveStyles();
  // Master Havuz — herkese açık, uygulama açılışında Supabase'den yüklenir
  const [pool,setPool]=useState([]);
  useEffect(()=>{
    supabase.from("pool").select("*").order("id").then(({data,error})=>{
      if(!error&&data)setPool(data.map(poolRowToJs));
    });
  },[]);

  // Eczane hesapları (süper admin için canlı liste — Supabase'den çekilir)
  const [accounts,setAccounts]=useState([]);
  // Giriş yapmış eczanenin kendi profili (accounts listesinden bağımsız — RLS gereği eczane sadece kendi satırını görebilir)
  const [loggedInAccount,setLoggedInAccount]=useState(null);
  const [authReady,setAuthReady]=useState(false);
  const [superadminEmail,setSuperadminEmail]=useState("");
  // Kiosk linkinden gelen eczane kimliği (?store=...) — girişsiz, herkese açık kiosk görünümü için
  const [publicStoreId]=useState(()=>getStoreParam());
  // O an gösterilecek kataloğun sahibi: önce giriş yapmış eczane, yoksa kiosk linkindeki eczane
  const activePharmacyId=loggedInAccount?.id||publicStoreId||null;

  // Eczane kataloğu + ayarları — aktif eczane belirlenince yüklenir, değişince (sadece giriş yapılmışsa) otomatik kaydedilir
  const [pharmacyCatalog,setPharmacyCatalog]=useState([]);
  const [pharmacySettings,setPharmacySettings]=useState({name:"Yeni Eczane",primaryBrand:"Nötr"});
  const [catalogReady,setCatalogReady]=useState(false);

  const syncCatalogToSupabase=async(pharmacyId,catalog)=>{
    await supabase.from("pharmacy_catalog").delete().eq("pharmacy_id",pharmacyId);
    if(catalog.length===0)return;
    const rows=catalog.map(p=>({
      pharmacy_id:pharmacyId, product_id:p.id,
      price:p.price||p.basePrice||0, discounted_price:p.discountedPrice||null,
      stock:typeof p.stock==="number"?p.stock:null,
      cost:p.cost||null, weekly_sales:p.weeklySales??null, skt:p.skt||null,
    }));
    await supabase.from("pharmacy_catalog").insert(rows);
  };
  const syncSettingsToSupabase=async(pharmacyId,settings)=>{
    await supabase.from("pharmacy_settings").update({
      name:settings.name, primary_brand:settings.primaryBrand,
      cart_discount_enabled:settings.cartDiscountEnabled!==false,
      cart_discount_threshold:settings.cartDiscountThreshold||1,
      cart_discount_pct:settings.cartDiscountPct||10,
    }).eq("pharmacy_id",pharmacyId);
  };

  // Aktif eczane değiştikçe (giriş / kiosk linki): kataloğu + ayarları yükle
  // (yoksa, sadece giriş yapılmışsa, varsayılan ayar satırı oluşturulur)
  useEffect(()=>{
    setCatalogReady(false);
    if(!activePharmacyId){
      setPharmacyCatalog([]);
      setPharmacySettings({name:"Yeni Eczane",primaryBrand:"Nötr"});
      return;
    }
    (async()=>{
      const[{data:catRows},{data:setRow}]=await Promise.all([
        supabase.from("pharmacy_catalog").select("*, pool(*)").eq("pharmacy_id",activePharmacyId),
        supabase.from("pharmacy_settings").select("*").eq("pharmacy_id",activePharmacyId).maybeSingle(),
      ]);
      setPharmacyCatalog((catRows||[]).filter(r=>r.pool).map(catalogRowToJs));
      if(setRow){
        setPharmacySettings({
          name:setRow.name, primaryBrand:setRow.primary_brand,
          cartDiscountEnabled:setRow.cart_discount_enabled,
          cartDiscountThreshold:setRow.cart_discount_threshold,
          cartDiscountPct:setRow.cart_discount_pct,
        });
      } else if(loggedInAccount?.id===activePharmacyId){
        // Sadece giriş yapmış eczanenin kendi ilk kaydı için varsayılan ayar satırı oluşturulur
        await supabase.from("pharmacy_settings").insert({pharmacy_id:activePharmacyId,name:"Yeni Eczane",primary_brand:"Nötr"});
        setPharmacySettings({name:"Yeni Eczane",primaryBrand:"Nötr"});
      }
      setCatalogReady(true);
    })();
  },[activePharmacyId]);

  // Katalog değişince (eşik geçtikten sonra) otomatik Supabase'e kaydet — sadece giriş yapmış eczane için
  useEffect(()=>{
    if(!catalogReady||!loggedInAccount?.id)return;
    const t=setTimeout(()=>{syncCatalogToSupabase(loggedInAccount.id,pharmacyCatalog);},900);
    return()=>clearTimeout(t);
  },[pharmacyCatalog,catalogReady,loggedInAccount?.id]);

  // Ayarlar değişince otomatik Supabase'e kaydet — sadece giriş yapmış eczane için
  useEffect(()=>{
    if(!catalogReady||!loggedInAccount?.id)return;
    const t=setTimeout(()=>{syncSettingsToSupabase(loggedInAccount.id,pharmacySettings);},900);
    return()=>clearTimeout(t);
  },[pharmacySettings,catalogReady,loggedInAccount?.id]);

  // Kiosk üzerinden tamamlanan satış/liste kayıtları (satış grafiği ve fayda analizi için)
  const [sales,setSales]=useState([]);
  const [salesReady,setSalesReady]=useState(false);

  const salesRowToJs=(r)=>({
    id:r.id, date:r.created_at, items:r.items||[],
    total:r.total, totalOrig:r.total_orig, itemCount:r.item_count,
    discountApplied:r.discount_applied,
  });

  const recordSale=async(sale)=>{
    if(!activePharmacyId)return; // hangi eczaneye ait olduğu bilinmiyorsa kaydedilmez
    const localId=Date.now()+Math.random();
    setSales(prev=>[...prev,{id:localId,...sale}]);

    // Satılan ürünlerin sayısal stok adedi varsa düş (Excel/manuel girilen stoklarla tutarlı kalsın)
    const soldIds=new Set((sale.items||[]).map(it=>it.id));
    if(soldIds.size>0){
      setPharmacyCatalog(prev=>prev.map(p=>{
        if(!soldIds.has(p.id)||typeof p.stock!=="number")return p;
        const qty=(sale.items||[]).filter(it=>it.id===p.id).reduce((s,it)=>s+(it.qty||1),0);
        return {...p,stock:Math.max(0,p.stock-qty)};
      }));
    }

    // Supabase'e kalıcı olarak kaydet (kiosk girişsiz/anonim olsa bile, onaylı eczaneye satış eklenebilir)
    await supabase.from("sales").insert({
      pharmacy_id:activePharmacyId, items:sale.items||[],
      total:sale.total||0, total_orig:sale.totalOrig||0,
      item_count:sale.itemCount||0, discount_applied:!!sale.discountApplied,
    });
  };

  // Giriş yapmış eczanenin kendi satış geçmişini yükle (analiz paneli için)
  useEffect(()=>{
    setSalesReady(false);
    if(!loggedInAccount?.id){setSales([]);setSalesReady(true);return;}
    supabase.from("sales").select("*").eq("pharmacy_id",loggedInAccount.id)
      .order("created_at",{ascending:true}).then(({data,error})=>{
        if(!error&&data)setSales(data.map(salesRowToJs));
        setSalesReady(true);
      });
  },[loggedInAccount?.id]);

  const resetSales=async()=>{
    if(!loggedInAccount?.id)return;
    setSales([]);
    await supabase.from("sales").delete().eq("pharmacy_id",loggedInAccount.id);
  };

  const urlMode=getUrlMode();
  const [screen,setScreen]=useState(()=>{
    if(urlMode==="pharmacy")return"pharmacy_auth";
    if(urlMode==="superadmin")return"superadmin_auth";
    if(urlMode==="admin")return"admin_select";
    return"customer";
  });

  const rowToAccount=(r)=>({
    id:r.id, pharmacyName:r.pharmacy_name||"", contactName:r.contact_name||"",
    email:r.email||"", phone:r.phone||"", status:r.status, createdAt:r.created_at,
  });

  const loadAccounts=async()=>{
    const{data,error}=await supabase.from("profiles").select("*")
      .eq("role","pharmacy").order("created_at",{ascending:false});
    if(!error&&data)setAccounts(data.map(rowToAccount));
  };

  // Sayfa açıldığında/yenilendiğinde mevcut Supabase oturumunu geri yükle
  useEffect(()=>{
    supabase.auth.getSession().then(async({data})=>{
      const user=data?.session?.user;
      if(user){
        const{data:prof}=await supabase.from("profiles").select("*").eq("id",user.id).single();
        if(prof){
          if(prof.role==="pharmacy"&&prof.status==="approved"){
            setLoggedInAccount(rowToAccount(prof));
            setScreen("pharmacy");
          } else if(prof.role==="superadmin"&&prof.status==="approved"){
            await loadAccounts();
            setSuperadminEmail(prof.email||user.email||"");
            setScreen("superadmin");
          }
        }
      }
      setAuthReady(true);
    });
  },[]);

  // ── Kayıt talebi oluştur (eczane) — Supabase Auth hesabı + otomatik profil ──
  const registerAccount=async(data)=>{
    const{error}=await supabase.auth.signUp({
      email:data.email, password:data.password,
      options:{data:{pharmacy_name:data.pharmacyName.trim(),contact_name:data.contactName.trim(),phone:data.phone.trim()}}
    });
    if(error){
      const msg=(error.message||"").toLowerCase();
      if(msg.includes("already") || msg.includes("registered"))
        return{ok:false,error:"Bu e-posta ile zaten bir kayıt var."};
      return{ok:false,error:error.message||"Kayıt sırasında bir hata oluştu."};
    }
    await supabase.auth.signOut(); // onay bekleyen hesap otomatik oturum açmasın
    return{ok:true};
  };

  // ── Eczane girişi: yalnızca onaylı hesaplar geçebilir ──
  const loginPharmacy=async(email,password)=>{
    const{data,error}=await supabase.auth.signInWithPassword({email,password});
    if(error||!data?.session){
      const msg=(error?.message||"").toLowerCase();
      if(msg.includes("confirm"))return{ok:false,error:"E-posta adresinizi onaylamanız gerekiyor. Gelen kutunuzu kontrol edin."};
      return{ok:false,error:"E-posta veya şifre hatalı."};
    }
    const{data:prof,error:profErr}=await supabase.from("profiles").select("*").eq("id",data.session.user.id).single();
    if(profErr||!prof){await supabase.auth.signOut();return{ok:false,error:"Hesap profili bulunamadı."};}
    if(prof.role!=="pharmacy"){await supabase.auth.signOut();return{ok:false,error:"Bu hesap eczane hesabı değil."};}
    if(prof.status==="pending"){await supabase.auth.signOut();return{ok:false,error:"Hesabınız henüz onaylanmadı. Lütfen süper admin onayını bekleyin."};}
    if(prof.status==="rejected"){await supabase.auth.signOut();return{ok:false,error:"Kayıt talebiniz reddedildi. Eczane yönetimiyle iletişime geçin."};}
    if(prof.status==="suspended"){await supabase.auth.signOut();return{ok:false,error:"Hesabınız askıya alındı. Yönetici ile iletişime geçin."};}
    const account=rowToAccount(prof);
    setLoggedInAccount(account);
    return{ok:true,account};
  };

  const loginSuperadmin=async(email,password)=>{
    const{data,error}=await supabase.auth.signInWithPassword({email,password});
    if(error||!data?.session)return{ok:false,error:"E-posta veya şifre hatalı."};
    const{data:prof,error:profErr}=await supabase.from("profiles").select("*").eq("id",data.session.user.id).single();
    if(profErr||!prof||prof.role!=="superadmin"||prof.status!=="approved"){
      await supabase.auth.signOut();
      return{ok:false,error:"Bu hesabın süper admin yetkisi yok."};
    }
    await loadAccounts();
    return{ok:true};
  };

  // ── Eczane kendi bilgilerini günceller (mevcut şifre yeniden giriş ile doğrulanır) ──
  const updateAccountSelf=async(id,{currentPassword,email,phone,newPassword})=>{
    const{data:{user}}=await supabase.auth.getUser();
    if(!user)return{ok:false,error:"Oturum bulunamadı, tekrar giriş yapın."};
    const{error:reauthErr}=await supabase.auth.signInWithPassword({email:user.email,password:currentPassword});
    if(reauthErr)return{ok:false,error:"Mevcut şifre yanlış."};

    const authUpdates={};
    if(email&&email!==user.email)authUpdates.email=email;
    if(newPassword)authUpdates.password=newPassword;
    if(Object.keys(authUpdates).length>0){
      const{error}=await supabase.auth.updateUser(authUpdates);
      if(error)return{ok:false,error:error.message};
    }
    const{error:profErr}=await supabase.from("profiles").update({phone,email}).eq("id",id);
    if(profErr)return{ok:false,error:profErr.message};
    setLoggedInAccount(prev=>prev?{...prev,email,phone}:prev);
    return{ok:true, notice: authUpdates.email?"E-posta değişikliğini onaylamak için yeni adresinize gelen bağlantıya tıklamanız gerekebilir.":undefined};
  };

  const updateSuperadminSelf=async({currentPassword,email,newPassword})=>{
    const{data:{user}}=await supabase.auth.getUser();
    if(!user)return{ok:false,error:"Oturum bulunamadı, tekrar giriş yapın."};
    const{error:reauthErr}=await supabase.auth.signInWithPassword({email:user.email,password:currentPassword});
    if(reauthErr)return{ok:false,error:"Mevcut şifre yanlış."};
    const authUpdates={};
    if(email&&email!==user.email)authUpdates.email=email;
    if(newPassword)authUpdates.password=newPassword;
    if(Object.keys(authUpdates).length===0)return{ok:true};
    const{error}=await supabase.auth.updateUser(authUpdates);
    if(error)return{ok:false,error:error.message};
    await supabase.from("profiles").update({email}).eq("id",user.id);
    return{ok:true};
  };

  // ── Süper admin hesap yönetimi ──
  const approveAccount=async(id)=>{await supabase.from("profiles").update({status:"approved"}).eq("id",id);await loadAccounts();};
  const rejectAccount=async(id)=>{await supabase.from("profiles").update({status:"rejected"}).eq("id",id);await loadAccounts();};
  const suspendAccount=async(id)=>{await supabase.from("profiles").update({status:"suspended"}).eq("id",id);await loadAccounts();};
  const reactivateAccount=async(id)=>{await supabase.from("profiles").update({status:"approved"}).eq("id",id);await loadAccounts();};
  const deleteAccount=async(id)=>{await supabase.from("profiles").delete().eq("id",id);await loadAccounts();};
  // Not: publishable key ile başka bir kullanıcının şifresi DOĞRUDAN değiştirilemez (güvenlik gereği,
  // bu servis rolü/sunucu tarafı yetkisi gerektirir). Bunun yerine kullanıcıya sıfırlama e-postası gönderilir.
  const resetAccountPassword=async(id)=>{
    const acc=accounts.find(a=>a.id===id);
    if(!acc)return;
    await supabase.auth.resetPasswordForEmail(acc.email);
  };

  const logoutPharmacy=async()=>{await supabase.auth.signOut();setLoggedInAccount(null);setScreen("admin_select");};
  const logoutSuperadmin=async()=>{await supabase.auth.signOut();setAccounts([]);setScreen("admin_select");};

  if(!authReady) return(
    <div style={{minHeight:"100vh",display:"flex",flexDirection:"column",gap:12,alignItems:"center",justifyContent:"center",
      background:"#F5F5F5",fontFamily:"'Inter',sans-serif",color:"#888",fontSize:13}}>
      <Spinner/>
      <span>Yükleniyor…</span>
    </div>
  );

  if(screen==="customer") return <CustomerKiosk
    pharmacyCatalog={pharmacyCatalog} pharmacySettings={pharmacySettings}
    onAdminTrigger={()=>setScreen("admin_select")} onCompleteSale={recordSale}/>;

  if(screen==="admin_select") return <AdminSelect
    onSelect={r=>setScreen(r+"_auth")} onBack={()=>setScreen("customer")}/>;

  if(screen==="pharmacy_auth") return <PharmacyAuthGate
    onLogin={loginPharmacy} onRegister={registerAccount}
    onSuccess={acc=>{setLoggedInAccount(acc);setScreen("pharmacy");}}
    onBack={()=>setScreen("admin_select")}/>;

  if(screen==="superadmin_auth") return <SuperAdminAuthGate
    onLogin={async(email,password)=>{
      const res=await loginSuperadmin(email,password);
      if(res.ok)setSuperadminEmail(email);
      return res;
    }}
    onSuccess={()=>setScreen("superadmin")} onBack={()=>setScreen("admin_select")}/>;

  if(screen==="pharmacy"){
    if(!loggedInAccount||loggedInAccount.status!=="approved"){
      // Hesap silinmiş/askıya alınmışsa oturumu düşür
      return <PharmacyAuthGate onLogin={loginPharmacy} onRegister={registerAccount}
        onSuccess={acc=>{setLoggedInAccount(acc);setScreen("pharmacy");}}
        onBack={()=>setScreen("admin_select")}/>;
    }
    return <PharmacyAdmin
      masterPool={pool} setPool={setPool} pharmacyCatalog={pharmacyCatalog} setPharmacyCatalog={setPharmacyCatalog}
      pharmacySettings={pharmacySettings} setPharmacySettings={setPharmacySettings}
      catalogReady={catalogReady}
      sales={sales} onResetSales={resetSales}
      account={loggedInAccount} onUpdateAccount={updateAccountSelf}
      onBack={logoutPharmacy}/>;
  }

  if(screen==="superadmin") return <SuperAdmin
    pool={pool} setPool={setPool}
    accounts={accounts} onApprove={approveAccount} onReject={rejectAccount}
    onSuspend={suspendAccount} onReactivate={reactivateAccount} onDeleteAccount={deleteAccount}
    onResetPassword={resetAccountPassword}
    superadminAccount={{email:superadminEmail}} onUpdateSuperadmin={updateSuperadminSelf}
    onBack={logoutSuperadmin}/>;

  return null;
}
