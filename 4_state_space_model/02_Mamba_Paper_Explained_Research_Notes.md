# Mamba: Linear-Time Sequence Modeling with Selective State Spaces (2024)

## Paper Explained (Research Notes)

> Based on the highlighted sections of the original paper. These notes
> are intended for research and understanding rather than presentation.

------------------------------------------------------------------------

# Preface

This document explains the core ideas behind **Mamba** from an intuitive
and research-oriented perspective. Rather than following the paper
paragraph by paragraph, it focuses on answering four questions
throughout:

1.  **What problem are the authors trying to solve?**
2.  **Why do previous approaches fail?**
3.  **How does Mamba solve the problem?**
4.  **Why is this important for future architectures?**

The emphasis is on language modeling. DNA and audio experiments are only
briefly discussed.

------------------------------------------------------------------------

# 1. Introduction

## Motivation

Transformers dominate modern foundation models because self-attention
performs **content-based reasoning** extremely well. However, this
capability comes with two major drawbacks:

-   Quadratic complexity with respect to sequence length.
-   Limited context window during inference.

Many alternatives (linear attention, recurrent models, SSMs) improve
efficiency but consistently underperform Transformers on language.

The authors argue that **efficiency was never the true limitation** of
previous SSMs. Instead, they lacked a mechanism to **select information
based on content**.

This observation becomes the foundation of Mamba.

### Main Contributions

1.  Introduce **Selective State Space Models (S6)** by making SSM
    parameters depend on the input.
2.  Design a **hardware-aware parallel scan** to keep computation
    efficient despite losing convolution.
3.  Build a simplified backbone (Mamba) entirely around selective SSM
    blocks.
4.  Demonstrate Transformer-level language modeling while retaining
    linear complexity.

------------------------------------------------------------------------

# 2. State Space Models (Recap)

A structured SSM models a sequence using a hidden state.

Instead of storing all previous tokens (as attention does), the sequence
is compressed into a latent state that evolves over time.

After discretization, an SSM can be computed in two equivalent forms:

-   **Recurrence**
-   **Convolution**

This equivalence is possible because classical SSMs are **Linear Time
Invariant (LTI)**.

## Why LTI Matters

LTI means the parameters governing state evolution remain constant for
every token.

Consequences:

-   Extremely efficient.
-   Convolution becomes possible.
-   But every token is processed using exactly the same dynamics.

This assumption ultimately limits language modeling.

------------------------------------------------------------------------

# 3. Why Classical SSMs Fail

The paper reframes sequence modeling as a compression problem.

Attention stores almost everything.

Recurrent models compress everything into a finite state.

Therefore the real question becomes:

> Can the model compress information without forgetting what actually
> matters?

Classical SSMs cannot.

Since their dynamics never change, they cannot decide whether a token
should be remembered or ignored.

This is exactly what language requires.

------------------------------------------------------------------------

# 4. Selection Mechanism

The central idea of Mamba is remarkably simple.

Instead of using fixed transition dynamics,

the transition becomes **input dependent**.

Rather than asking

> "How should memory evolve?"

Mamba asks

> "Given this token, how should memory evolve?"

To accomplish this,

-   Δ becomes input dependent.
-   B becomes input dependent.
-   C becomes input dependent.

This transforms the model from **time invariant** into **time varying**.

The consequence is enormous:

The model can now

-   remember,
-   forget,
-   overwrite,
-   ignore

information according to the current token.

This is the key innovation of the paper.

Notice that A remains fixed to preserve the mathematical structure and efficiency of the SSM.

------------------------------------------------------------------------

# 5. Why Synthetic Tasks?

The synthetic tasks are not benchmarks.

They are diagnostic experiments.

## Selective Copying

The classical Copying task only tests memory.

An LTI model can solve it simply by learning the correct temporal
spacing.

Selective Copying randomizes token positions.

Now the model must identify **which tokens matter**, not merely **when**
they occur.

This isolates content-aware reasoning.

## Induction Heads

This task evaluates associative recall.

If the model previously observed

Harry → Potter

then after seeing "Harry" again it should retrieve "Potter".

This ability is believed to underlie in-context learning in LLMs.

Mamba solves both tasks because it selectively stores relevant
information while filtering irrelevant context.

------------------------------------------------------------------------

# 6. Hardware-Aware Selective Scan

Making parameters input-dependent breaks convolution.

The challenge becomes:

How can a time-varying recurrence remain efficient?

The solution is the Selective Scan algorithm.

Instead of materializing the enormous hidden state in GPU HBM,

the algorithm

-   loads parameters into SRAM,
-   discretizes there,
-   performs recurrence there,
-   writes only the final output back.

Additional optimizations include

-   kernel fusion,
-   parallel scan,
-   recomputation.

Together they preserve linear complexity while avoiding excessive memory
traffic.

------------------------------------------------------------------------

# 7. Mamba Architecture

Previous SSM architectures alternated

SSM → MLP → SSM → MLP

Mamba merges these ideas into one homogeneous block.

Compared with H3:

-   simpler,
-   fewer architectural components,
-   easier to scale.

Most parameters remain inside linear projections.

The SSM contributes relatively few parameters.

------------------------------------------------------------------------

# 8. Why Selection Works

Selection introduces several important properties.

## Variable Spacing

Noise tokens can simply be ignored.

## Filtering Context

The model can reset memory when context becomes irrelevant.

Therefore longer context generally helps instead of hurting.

## Boundary Resetting

Independent sequences can be separated without explicit attention masks.

------------------------------------------------------------------------

# 9. Connection with RNN Gates

One elegant contribution of the paper is showing that

classical RNN gating

is a special case of selective SSMs.

In particular,

Δ plays the role of a generalized gate.

Large Δ

-   resets memory,
-   emphasizes current input.

Small Δ

-   preserves memory,
-   ignores transient inputs.

Thus discretization provides a principled interpretation of heuristic
gating mechanisms.

------------------------------------------------------------------------

# 10. Language Modeling

This is the most important experimental section.

## Scaling Laws

Scaling laws study how performance changes as computational resources
increase.

Performance is measured using **Perplexity (PPL)**.

Perplexity measures how well the model predicts the next token.

Lower perplexity indicates better language modeling.

The results show that Mamba scales more efficiently than previous
attention-free architectures.

Most importantly,

Mamba becomes the first attention-free model capable of matching a
modern Transformer++ recipe.

The advantage becomes larger as context length increases.

## Downstream Evaluation

The paper evaluates Mamba against models such as Pythia and RWKV on
standard zero-shot language understanding tasks.

Across model sizes,

Mamba consistently achieves the best overall performance among models
trained under comparable settings.

In many cases,

Mamba matches or exceeds Transformer baselines roughly twice its size.

------------------------------------------------------------------------

# 11. DNA and Audio

The paper also validates Mamba on genomics and audio.

These experiments serve one main purpose:

to demonstrate that Mamba is not merely a language model architecture
but a **general sequence modeling backbone**.

In both domains,

longer context improves performance,

supporting the authors' claim that selective memory benefits many
sequential modalities.

------------------------------------------------------------------------

# Final Takeaways

The contribution of Mamba is **not** simply replacing attention.

Its true contribution is introducing **content-aware state transitions**
into structured state space models while preserving linear-time
computation.

The paper shows that efficiency alone is insufficient for language
modeling.

What sequence models require is **selective memory**: the ability to
decide, at every token, what should be remembered and what should be
forgotten.

This single idea transforms SSMs from efficient recurrent models into
competitive general-purpose foundation model backbones.

------------------------------------------------------------------------

# From S4 to Mamba: Why Selection Matters

The key limitation of classical Structured State Space Models (SSMs), including S4, is that they are **Linear Time-Invariant (LTI)** systems.

This means that every token is processed using exactly the same transition dynamics.

Mathematically,

$$
A,\;B,\;C,\;\Delta
$$

remain fixed throughout the entire sequence.

Although this property enables an efficient convolutional implementation, it also means that the model has **no mechanism to distinguish important tokens from irrelevant ones**.

Every input influences memory in essentially the same way.

For language, this assumption is too restrictive.

Natural language contains highly informative tokens (names, entities, verbs) mixed with many low-information tokens (articles, punctuation, common words).

An effective sequence model should not treat all of them equally.

Instead, it should be able to answer questions such as:

- Should this token be stored?
- Should previous memory be overwritten?
- Can this token be ignored?

S4 cannot answer these questions because its dynamics never change.

---

## Mamba's Main Contribution: Selection

Mamba removes the Linear Time-Invariant assumption.

Instead of using fixed parameters,

$$
B,\;C,\;\Delta
$$

become functions of the current input,

$$
B(x_t),\qquad C(x_t),\qquad \Delta(x_t).
$$

As a consequence, the evolution of the hidden state depends not only on time, but also on **what the current token actually is**.

This mechanism is called **selection**.

Rather than processing every token identically, the model dynamically decides:

- what information should enter memory,
- what information should be forgotten,
- which context should be preserved,
- and which tokens can safely be ignored.

The hidden state therefore becomes **content-aware** rather than purely time-dependent.

---

## Intuition

Suppose the sequence is

the  the  the  Harry  Potter  .  the

A classical S4 processes every token with the same transition dynamics.

Mamba, however, can implicitly behave as

ignore → ignore → ignore → store → store → reset → ignore

because each token generates different values of

$$
B(x),\;C(x),\;\Delta(x).
$$

The model is no longer forced to compress every token equally.

Instead, it learns **what is worth remembering**.

This single idea transforms SSMs from efficient recurrent models into competitive language models.

It is the central innovation of the Mamba paper.