// Builds the ATS-friendly CVs (single column, standard fonts and headings).
// Usage: NODE_PATH=$(npm root -g) node tools/build-cv-docx.js  (needs: npm i -g docx)   →  Renato-Garcia-CV-EN.docx, Renato-Garcia-Curriculo-PT.docx
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, LevelFormat, BorderStyle,
  TabStopType, ExternalHyperlink,
} = require('docx');

const CONTACT = 'Londrina, PR, Brazil · +55 43 99639-9184 · renato.java@gmail.com';
const LINKS = [
  ['linkedin.com/in/renato-garcia-java', 'https://www.linkedin.com/in/renato-garcia-java/'],
  ['renatojava.github.io', 'https://renatojava.github.io/'],
  ['github.com/renatojava', 'https://github.com/renatojava'],
];

const CV = {
  en: {
    file: 'Renato-Garcia-CV-EN.docx',
    title: 'Software Architect | Distributed Systems & Applied AI | Targeting Staff / Principal Engineer roles',
    h: { summary: 'Professional Summary', skills: 'Skills', exp: 'Experience', certs: 'Certifications', edu: 'Education', lang: 'Languages' },
    summary: [
      'Software architect with 20+ years building distributed systems, most of them integration platforms that move data reliably at volume. Architect of GuBee since day one: an event-driven hub on Apache Kafka connecting e-commerce platforms to 30+ Brazilian marketplaces, today about 170 production workloads. Founding engineer of INAI (Integra.do), a multi-agent LLM platform live in production since January 2026.',
      'I set technical direction across teams (architecture standards, cloud and platform migrations, roadmap), lead AI products from framework choice to production, and stay hands-on in Java and Kotlin.',
    ],
    skills: [
      ['Architecture', 'Distributed Systems, Microservices, Event-Driven Architecture, Domain-Driven Design (DDD), Hexagonal Architecture, TDD, System Design, Enterprise Integration Patterns'],
      ['Languages & frameworks', 'Java 25, Kotlin, Spring Boot, Quarkus (native / GraalVM), Spring AI, JetBrains Koog, REST APIs'],
      ['Applied AI', 'LLM Applications, Multi-Agent Systems, RAG, pgvector, OpenAI, MCP, Guardrails, LLM Evaluation (DeepEval), LLM Observability'],
      ['Data & messaging', 'Apache Kafka, Redis / Redis Streams, MongoDB, PostgreSQL / PostGIS, OpenSearch, AWS SQS/SNS, Oracle'],
      ['Cloud & delivery', 'Kubernetes, Docker, AWS (EKS, Lambda), Tencent Cloud, GitLab CI/CD, OpenTelemetry, Grafana (Tempo, Loki, Mimir), Prometheus'],
      ['Leadership', 'Technical Leadership, Mentoring, Technical Roadmap, Cloud Migration, Stakeholder Management'],
    ],
    jobs: [
      { role: 'Founding Engineer & AI Architect', org: 'Integra.do (INAI)', meta: 'Remote', when: 'Jul 2025 – Present', bullets: [
        'Architected and built INAI from the first commit: a multi-agent LLM platform that qualifies and serves real-estate leads 24/7 on WhatsApp, live in production since January 2026 at Imobiliária Inglaterra (launch client).',
        'Designed the multi-agent core: 5 specialist agents (SDR, billing, support, maintenance, launches) behind a hybrid router (keyword + semantic + session affinity, LLM tie-break) with agent-to-agent handoff and per-tenant overrides.',
        'Built zero-loss WhatsApp ingestion on Redis Streams (consumer groups, pending-entry recovery, per-tenant debounce) and lead ingestion from ZAP, VivaReal, OLX and ImovelWeb with deduplication and replay protection.',
        'Implemented retrieval and agent runtime: hybrid search (pgvector + SQL + full text, RRF) with a local ONNX reranker, dynamic tools (HTTP, sandboxed GraalJS, MCP), three-layer memory, semantic LLM cache and two-tier guardrails.',
        'Controlled LLM cost with four model tiers by task and 11 Grafana dashboards tracking cost, latency and prompt-cache hit ratio per tenant (OpenTelemetry, Prometheus).',
        'Led a hardening milestone: tenant isolation on 43 controllers, BCrypt API keys migrated with a dual-read window, ~200 blocking calls removed; ~3,200 automated tests.',
        'Wrote ~70% of 2,400+ commits and every architecture and roadmap document (multi-tenant SaaS, admin/RBAC, CRM connectors, SSO, SDKs).',
      ], tech: 'Kotlin, Java 25, Spring Boot 4, Spring AI, OpenAI, PostgreSQL, pgvector, PostGIS, Redis Streams, Kubernetes, GitLab CI, React' },
      { role: 'Software Architect', org: 'GuBee', meta: 'Full-time · Londrina, PR', when: 'Jul 2019 – Present', bullets: [
        'Own the architecture of an integration hub connecting e-commerce platforms (VTEX, Shopify, Magento, WooCommerce) to 30+ marketplaces (Mercado Livre, Amazon, Magalu, Shopee): about 170 production workloads on an event-driven Apache Kafka backbone with Redis and MongoDB.',
        'Planned and lead the cloud migration from Huawei Cloud to Tencent Cloud: ~170 workloads, 36+ MongoDB databases and 25+ Kafka consumer groups, cut over behind feature flags with rollback at every step; wrote the playbooks that let several engineers migrate services in parallel.',
        'Drive fleet-wide modernisation: Java 25, Quarkus native with GraalVM (over 60% less memory per service), and OpenTelemetry traces, metrics and logs on Grafana Tempo, Loki and Mimir.',
        'Lead GuBee’s AI products on JetBrains Koog, chosen after head-to-head POCs against Embabel, LangChain4j and Spring AI: conversational commerce on WhatsApp (discovery to PIX checkout) and AI catalog enrichment with human-in-the-loop; quality gated by DeepEval (faithfulness, refusal, prompt injection).',
        'Built the company’s AI engineering toolkit: 15 Claude Code plugins for migrations, scaffolding, troubleshooting and architecture checks, so every team applies the same standards.',
        'Grew and lead the engineering team; set Hexagonal Architecture, DDD and TDD as the default.',
      ], tech: 'Java 25, Kotlin, Quarkus, Spring Boot, Apache Kafka, Redis, MongoDB, OpenSearch, Kubernetes, AWS, Tencent Cloud, OpenTelemetry, Koog' },
      { role: 'Tech Lead', org: 'Serasa Experian', meta: 'Contract · Remote', when: 'Jan 2025 – May 2025', bullets: [
        'Technical lead of the biometrics platform team: defined the architecture and owned delivery quality.',
        'Ran technical alignment and requirements refinement with product and business stakeholders.',
        'Built microservices and asynchronous flows on AWS SQS/SNS for biometric data processing.',
      ], tech: 'Java, Spring Boot, AWS SQS/SNS, Oracle Autonomous Database' },
      { role: 'System Architect', org: 'Atos', meta: 'Londrina / São Paulo', when: 'Dec 2014 – Jul 2019', bullets: [
        'Architect on enterprise integration programmes for NET/Claro, Multiplus Fidelidade and Carrefour.',
        'Designed OSB, BPEL and BPM interfaces on Oracle SOA Suite following Enterprise Integration Patterns.',
        'Led the product definition for Carrefour’s HeavyGoods platform, built on event-based microservices.',
      ], tech: 'Java, Oracle SOA Suite, OSB, BPEL, Microservices' },
      { role: 'Software Architect & Project Lead', org: 'Infracommerce', meta: 'Consultant', when: 'Nov 2013 – Nov 2014', bullets: [
        'Led projects on the ACEC e-commerce platform for Carrefour Group, Decathlon and Lança Perfume.',
        'Built a service-oriented architecture with administration, sales and ERP integration layers; Drools for dynamic pricing, Solr for search.',
      ], tech: 'Java, Oracle ADF, JSF, Solr, Drools, BPEL, OSB' },
      { role: 'Tech Lead', org: 'Accurate', meta: 'Londrina / São Paulo', when: 'Aug 2010 – Nov 2014', bullets: [
        'Team lead on CPQD public-sector projects: requirements, function-point estimation, architecture and team coordination.',
      ] },
      { role: 'System Architect', org: 'Atos (client: Serasa)', meta: 'São Paulo', when: 'Jun 2013 – Nov 2013', bullets: [
        'Architecture definition and technical specification for credit applications.',
      ] },
      { role: 'Tech Lead', org: 'LINT', meta: 'Londrina', when: 'Jun 2009 – Nov 2010', bullets: [
        'Led development of a city administration system for CPQD, end to end.',
      ] },
      { role: 'Systems Analyst', org: 'Amadeus IT Group', meta: '', when: '2008', bullets: [
        'Insurance applications and a tourism platform for Carrefour (Java EE, Spring, JPA, web services).',
      ] },
      { role: 'Team Leader & Quality Assurance', org: 'BSI Tecnologia', meta: 'São Paulo', when: '2003 – 2008', bullets: [
        'Grew from developer to team leader; banking projects for Itaú and Bradesco integrating Java with mainframe COBOL.',
      ] },
    ],
    certs: [
      'Sun Certified Enterprise Architect (2010)',
      'Oracle Certified Professional, Java SE 7 Programmer (2014)',
      'Professional Scrum Master I (2018)',
      'Oracle SOA Suite 11g Essentials (2015) · ITIL Foundation (2013)',
      'Sun Certified: Business Component Developer Java EE 5 and 1.3 (2007), Web Component Developer (2006), Programmer (2006), Associate (2008)',
    ],
    edu: [
      ['Postgraduate, Computer Software Engineering', 'Unifil — Centro Universitário Filadélfia', '2014 – 2015'],
      ['Bachelor of Science, Information Systems', 'Universidade Tecnológica Federal do Paraná (UTFPR)', '2003 – 2007'],
    ],
    lang: 'Portuguese: native · English: upper-intermediate (CEFR B2)',
  },
  pt: {
    file: 'Renato-Garcia-Curriculo-PT.docx',
    title: 'Arquiteto de Software | Sistemas Distribuídos e IA Aplicada | Foco em posições Staff / Principal Engineer',
    h: { summary: 'Resumo Profissional', skills: 'Competências', exp: 'Experiência Profissional', certs: 'Certificações', edu: 'Formação Acadêmica', lang: 'Idiomas' },
    summary: [
      'Arquiteto de software com mais de 20 anos construindo sistemas distribuídos, principalmente plataformas de integração que movimentam dados com confiabilidade e em alto volume. Arquiteto da GuBee desde o início: um hub orientado a eventos em Apache Kafka que conecta plataformas de e-commerce a mais de 30 marketplaces, hoje com cerca de 170 workloads em produção. Founding engineer do INAI (Integra.do), plataforma multiagente de LLM em produção desde janeiro de 2026.',
      'Defino a direção técnica entre times (padrões de arquitetura, migrações de nuvem e de plataforma, roadmap), lidero produtos de IA da escolha do framework até a produção e sigo hands-on em Java e Kotlin.',
    ],
    skills: [
      ['Arquitetura', 'Sistemas Distribuídos, Microsserviços, Arquitetura Orientada a Eventos, Domain-Driven Design (DDD), Arquitetura Hexagonal, TDD, System Design, Padrões de Integração (EIP)'],
      ['Linguagens e frameworks', 'Java 25, Kotlin, Spring Boot, Quarkus (native / GraalVM), Spring AI, JetBrains Koog, APIs REST'],
      ['IA aplicada', 'Aplicações com LLM, Sistemas Multiagente, RAG, pgvector, OpenAI, MCP, Guardrails, Avaliação de LLM (DeepEval), Observabilidade de LLM'],
      ['Dados e mensageria', 'Apache Kafka, Redis / Redis Streams, MongoDB, PostgreSQL / PostGIS, OpenSearch, AWS SQS/SNS, Oracle'],
      ['Cloud e entrega', 'Kubernetes, Docker, AWS (EKS, Lambda), Tencent Cloud, GitLab CI/CD, OpenTelemetry, Grafana (Tempo, Loki, Mimir), Prometheus'],
      ['Liderança', 'Liderança Técnica, Mentoria, Roadmap Técnico, Migração de Nuvem, Gestão de Stakeholders'],
    ],
    jobs: [
      { role: 'Founding Engineer e Arquiteto de IA', org: 'Integra.do (INAI)', meta: 'Remoto', when: 'jul 2025 – atual', bullets: [
        'Arquitetei e construí o INAI desde o primeiro commit: plataforma multiagente de LLM que qualifica e atende leads imobiliários 24/7 no WhatsApp, em produção desde janeiro de 2026 na Imobiliária Inglaterra (cliente de lançamento).',
        'Desenhei o núcleo multiagente: 5 agentes especialistas (SDR, cobrança, suporte, manutenção, lançamentos) com roteador híbrido (palavra-chave + semântica + afinidade de sessão, desempate por LLM), handoff entre agentes e overrides por cliente.',
        'Construí a ingestão de mensagens do WhatsApp sem perda em Redis Streams (consumer groups, recuperação de pendentes, debounce por cliente) e a captação de leads de ZAP, VivaReal, OLX e ImovelWeb com deduplicação e proteção contra replay.',
        'Implementei busca e runtime de agentes: busca híbrida (pgvector + SQL + texto, RRF) com reranker ONNX local, ferramentas dinâmicas (HTTP, GraalJS em sandbox, MCP), memória em três camadas, cache semântico de LLM e guardrails em dois níveis.',
        'Controlei o custo de LLM com quatro faixas de modelo por tarefa e 11 dashboards Grafana de custo, latência e taxa de acerto de prompt cache por cliente (OpenTelemetry, Prometheus).',
        'Conduzi um milestone de hardening: isolamento de tenant em 43 controllers, API keys com BCrypt migradas com leitura dupla, ~200 chamadas bloqueantes removidas; ~3.200 testes automatizados.',
        'Autor de ~70% dos 2.400+ commits e de todos os documentos de arquitetura e roadmap (SaaS multi-tenant, admin/RBAC, conectores de CRM, SSO, SDKs).',
      ], tech: 'Kotlin, Java 25, Spring Boot 4, Spring AI, OpenAI, PostgreSQL, pgvector, PostGIS, Redis Streams, Kubernetes, GitLab CI, React' },
      { role: 'Arquiteto de Software', org: 'GuBee', meta: 'CLT · Londrina, PR', when: 'jul 2019 – atual', bullets: [
        'Responsável pela arquitetura de um hub que integra plataformas de e-commerce (VTEX, Shopify, Magento, WooCommerce) a mais de 30 marketplaces (Mercado Livre, Amazon, Magalu, Shopee): cerca de 170 workloads em produção sobre Apache Kafka, com Redis e MongoDB.',
        'Planejei e lidero a migração de nuvem da Huawei Cloud para a Tencent Cloud: ~170 workloads, 36+ bancos MongoDB e 25+ consumer groups Kafka, com virada por feature flags e rollback em cada etapa; escrevi os playbooks que permitem a vários engenheiros migrar serviços em paralelo.',
        'Conduzo a modernização de toda a plataforma: Java 25, Quarkus native com GraalVM (mais de 60% menos memória por serviço) e OpenTelemetry (traces, métricas e logs) em Grafana Tempo, Loki e Mimir.',
        'Lidero os produtos de IA da GuBee com JetBrains Koog, escolhido após POCs comparativas com Embabel, LangChain4j e Spring AI: comércio conversacional no WhatsApp (da descoberta ao checkout com PIX) e enriquecimento de catálogo com IA e revisão humana; qualidade validada com DeepEval (fidelidade, recusa, prompt injection).',
        'Criei o toolkit de engenharia com IA da empresa: 15 plugins do Claude Code para migrações, scaffolding, troubleshooting e validação de arquitetura, para que todos os times sigam os mesmos padrões.',
        'Formei e lidero o time de engenharia; defini Arquitetura Hexagonal, DDD e TDD como padrão.',
      ], tech: 'Java 25, Kotlin, Quarkus, Spring Boot, Apache Kafka, Redis, MongoDB, OpenSearch, Kubernetes, AWS, Tencent Cloud, OpenTelemetry, Koog' },
      { role: 'Tech Lead', org: 'Serasa Experian', meta: 'Contrato · Remoto', when: 'jan 2025 – mai 2025', bullets: [
        'Liderança técnica do time da plataforma de biometria: definição de arquitetura e qualidade das entregas.',
        'Condução de alinhamentos técnicos e refinamento de requisitos com produto e negócio.',
        'Desenvolvimento de microsserviços e fluxos assíncronos em AWS SQS/SNS para processamento de dados biométricos.',
      ], tech: 'Java, Spring Boot, AWS SQS/SNS, Oracle Autonomous Database' },
      { role: 'Arquiteto de Sistemas', org: 'Atos', meta: 'Londrina / São Paulo', when: 'dez 2014 – jul 2019', bullets: [
        'Arquiteto em programas de integração corporativa para NET/Claro, Multiplus Fidelidade e Carrefour.',
        'Desenho de interfaces OSB, BPEL e BPM em Oracle SOA Suite seguindo Enterprise Integration Patterns.',
        'Liderei a definição do produto HeavyGoods do Carrefour, baseado em microsserviços orientados a eventos.',
      ], tech: 'Java, Oracle SOA Suite, OSB, BPEL, Microsserviços' },
      { role: 'Arquiteto de Software e Líder de Projetos', org: 'Infracommerce', meta: 'Consultor', when: 'nov 2013 – nov 2014', bullets: [
        'Liderei projetos na plataforma de e-commerce ACEC para Grupo Carrefour, Decathlon e Lança Perfume.',
        'Arquitetura orientada a serviços com camadas de administração, vendas e integração com ERP; Drools para precificação dinâmica e Solr para busca.',
      ], tech: 'Java, Oracle ADF, JSF, Solr, Drools, BPEL, OSB' },
      { role: 'Tech Lead', org: 'Accurate', meta: 'Londrina / São Paulo', when: 'ago 2010 – nov 2014', bullets: [
        'Líder técnico em projetos do CPQD para o setor público: requisitos, estimativa por pontos de função, arquitetura e coordenação do time.',
      ] },
      { role: 'Arquiteto de Sistemas', org: 'Atos (cliente: Serasa)', meta: 'São Paulo', when: 'jun 2013 – nov 2013', bullets: [
        'Definição de arquitetura e especificação técnica de aplicações de crédito.',
      ] },
      { role: 'Tech Lead', org: 'LINT', meta: 'Londrina', when: 'jun 2009 – nov 2010', bullets: [
        'Liderei de ponta a ponta o desenvolvimento de um sistema de gestão municipal para o CPQD.',
      ] },
      { role: 'Analista de Sistemas', org: 'Amadeus IT Group', meta: '', when: '2008', bullets: [
        'Aplicações de seguros e plataforma de turismo para o Carrefour (Java EE, Spring, JPA, web services).',
      ] },
      { role: 'Líder de Equipe e Qualidade', org: 'BSI Tecnologia', meta: 'São Paulo', when: '2003 – 2008', bullets: [
        'De desenvolvedor a líder de equipe; projetos bancários para Itaú e Bradesco integrando Java ao mainframe COBOL.',
      ] },
    ],
    certs: [
      'Sun Certified Enterprise Architect (2010)',
      'Oracle Certified Professional, Java SE 7 Programmer (2014)',
      'Professional Scrum Master I (2018)',
      'Oracle SOA Suite 11g Essentials (2015) · ITIL Foundation (2013)',
      'Sun Certified: Business Component Developer Java EE 5 e 1.3 (2007), Web Component Developer (2006), Programmer (2006), Associate (2008)',
    ],
    edu: [
      ['Pós-graduação em Engenharia de Software', 'Unifil — Centro Universitário Filadélfia', '2014 – 2015'],
      ['Bacharelado em Sistemas de Informação', 'Universidade Tecnológica Federal do Paraná (UTFPR)', '2003 – 2007'],
    ],
    lang: 'Português: nativo · Inglês: intermediário-avançado (CEFR B2)',
  },
};

const FONT = 'Calibri';
const RIGHT = 10178; // A4 text width (11906) minus 0.6" side margins, in DXA

const heading = (text) => new Paragraph({
  spacing: { before: 160, after: 60 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '3730A3', space: 2 } },
  children: [new TextRun({ text: text.toUpperCase(), bold: true, size: 22, color: '3730A3' })],
});
const para = (text, opts = {}) => new Paragraph({ spacing: { after: 80 }, ...opts, children: [new TextRun({ text, size: 20 })] });
const bullet = (text) => new Paragraph({ numbering: { reference: 'dot', level: 0 }, spacing: { after: 20 }, children: [new TextRun({ text, size: 20 })] });

function build(cv) {
  const out = [
    new Paragraph({ children: [new TextRun({ text: 'Renato Garcia', bold: true, size: 36 })] }),
    new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: cv.title, bold: true, size: 21, color: '3730A3' })] }),
    new Paragraph({ children: [new TextRun({ text: CONTACT, size: 19 })] }),
    new Paragraph({ spacing: { after: 60 }, children: LINKS.flatMap(([label, url], i) => [
      ...(i ? [new TextRun({ text: ' · ', size: 19 })] : []),
      new ExternalHyperlink({ link: url, children: [new TextRun({ text: label, size: 19, style: 'Hyperlink' })] }),
    ]) }),
    heading(cv.h.summary),
    ...cv.summary.map((t) => para(t)),
    heading(cv.h.skills),
    ...cv.skills.map(([k, v]) => new Paragraph({ spacing: { after: 40 }, children: [
      new TextRun({ text: `${k}: `, bold: true, size: 20 }), new TextRun({ text: v, size: 20 }),
    ] })),
    heading(cv.h.exp),
  ];
  for (const j of cv.jobs) {
    out.push(new Paragraph({
      spacing: { before: 120, after: 0 }, keepNext: true,
      tabStops: [{ type: TabStopType.RIGHT, position: RIGHT }],
      children: [
        new TextRun({ text: `${j.role} — ${j.org}`, bold: true, size: 21 }),
        new TextRun({ text: `\t${j.when}`, bold: true, size: 20 }),
      ],
    }));
    if (j.meta) out.push(new Paragraph({ keepNext: true, spacing: { after: 40 }, children: [new TextRun({ text: j.meta, italics: true, size: 19, color: '555555' })] }));
    j.bullets.forEach((b) => out.push(bullet(b)));
    if (j.tech) out.push(new Paragraph({ spacing: { after: 40 }, children: [
      new TextRun({ text: 'Tech: ', bold: true, size: 19 }), new TextRun({ text: j.tech, size: 19 }),
    ] }));
  }
  out.push(heading(cv.h.certs), ...cv.certs.map(bullet));
  out.push(heading(cv.h.edu), ...cv.edu.map(([deg, school, when]) => new Paragraph({
    spacing: { after: 40 }, tabStops: [{ type: TabStopType.RIGHT, position: RIGHT }],
    children: [new TextRun({ text: `${deg} — ${school}`, size: 20 }), new TextRun({ text: `\t${when}`, size: 20 })],
  })));
  out.push(heading(cv.h.lang), para(cv.lang));

  return new Document({
    creator: 'Renato Garcia',
    title: `Renato Garcia — ${cv.title}`,
    styles: { default: { document: { run: { font: FONT } } } },
    numbering: { config: [{ reference: 'dot', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 300, hanging: 220 } } } }] }] },
    sections: [{ properties: { page: { margin: { top: 600, bottom: 600, left: 864, right: 864 } } }, children: out }],
  });
}

(async () => {
  const root = path.join(__dirname, '..');
  for (const cv of Object.values(CV)) {
    fs.writeFileSync(path.join(root, cv.file), await Packer.toBuffer(build(cv)));
    console.log('wrote', cv.file);
  }
})();
