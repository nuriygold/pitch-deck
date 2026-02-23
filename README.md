# Nuriy Investor Deck  
## Verification Infrastructure for Supply Chain Compliance

This repository contains the official investor pitch deck for **Nuriy**, a compliance-grade verification infrastructure company building the product-level transparency standard for jewelry and adjacent consumer goods.

Nuriy is not a marketplace-first business.  
It is a **B2B compliance infrastructure layer** designed to help brands, retailers, procurement teams, and regulators verify supply chains at the product level.

---

## Company Overview

Nuriy is building what LEED did for buildings, but for jewelry and eventually for fashion, cosmetics, and food.

As global regulations such as the EU Corporate Sustainability Due Diligence Directive (CSDDD) move toward mandatory supply chain verification, brands will require independent, auditable, product-level transparency infrastructure.

Nuriy provides:

- A structured 0–100 transparency scoring methodology
- AI-powered audit synthesis and evidence classification
- Cryptographic integrity layer for tamper-evident verification
- API-based infrastructure for integration into brand and retail systems
- A certification and licensing model for brands to earn and display Nuriy Scores

Jewelry is the beachhead.  
Verification infrastructure is the long-term business.

---

## Repository Purpose

This repository generates and maintains the Nuriy investor deck programmatically using Node.js and PptxGenJS.

The goal is to:

- Maintain version control of the deck narrative
- Ensure structural alignment with compliance-first positioning
- Enable rapid iteration before investor conversations
- Keep messaging consistent with infrastructure and certification framing

---

## Deck Narrative Structure

The current deck follows this strategic arc:

1. Compliance-driven problem framing
2. Regulatory tailwinds (CSDDD, Modern Slavery Act, emerging US obligations)
3. Market opportunity as non-discretionary B2B spend
4. The Nuriy Standard (LEED analogy)
5. Proprietary infrastructure stack
6. B2B SaaS business model
7. Phased go-to-market strategy
8. Infrastructure-ready traction
9. Competitive infrastructure gap
10. Category expansion roadmap
11. MRR-focused financial model
12. Capital allocation for infrastructure activation

Marketplace mechanics are positioned only as a calibration and demonstration environment, not the core business.

---

## Technical Setup

### 1. Install Dependencies

```bash
npm install

Build the Deck
node build_deck.js
The generated file will output to:
/output/nuriy_compliance_infrastructure_deck.pptx
