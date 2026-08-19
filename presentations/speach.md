Good morning, and thank you for giving me the opportunity to present my doctoral research plan.

My name is Maximo Pacheco, and today I will discuss the plan entitled "Robust Temporal Representation Learning for Weakly Supervised Video Anomaly Detection."

==

First, let me briefly introduce the roadmap of today's presentation.

I will begin with my research background. Then, I will introduce the research context and explain the problem addressed by this PhD project.

After that, I will present the scientific gap, objective and the central hypothesis of this research. And finally, I will describe the proposed three-stage methodology, previous research problems, the evaluation strategy, and the expected contributions of this work.

==

Now, let's move on to my academic background.

I am from Mexico and earned both my bachelor's and master's degrees in Computer Science from the University of Puebla. Throughout my academic career, I have worked in several areas of Artificial Intelligence, particularly Natural Language Processing.

My previous research includes two peer-reviewed publications. The first focused on attribute selection using Binary Harmony Search, while the second explored hyperparameter optimization for a low-resource Nahuatl–Spanish Transformer-based machine translation system.

Regarding my academic training, my undergraduate thesis focused on recurrent neural networks for speech recognition and Spanish–Japanese machine translation. Later, during my master's degree, my research focused on developing a machine translation system from Nahuatl to Spanish.

In addition, I have received several academic scholarships and have presented my research at national conferences.

Overall, this background has provided me with a strong foundation in machine learning and sequence modeling, which naturally motivates my transition toward computer vision and video anomaly detection.

===

What is Video Anomaly Detection (VAD)?


Video Anomaly Detection (VAD) is a computer vision task that aims to identify events in a video that differ significantly from learned normal patterns. The challenge lies in the rarity and diversity of anomalous events, making it difficult to model all possible abnormal behaviors.

===

Why Video Anomaly Detection (VAD) Matters?

Well, because it has a wide range of real-world applications.

For example, in intelligent surveillance, anomaly detection can help automatically identify events such as accidents, robberies, or other unusual activities, improving public safety.

It is also relevant in industrial inspection, where it can detect abnormal machine behavior, and in autonomous systems, where recognizing unexpected events is essential for safe decision-making.

Despite its importance, Video Anomaly Detection still faces several research challenges.

The first challenge is learning effective temporal representations that accurately capture the temporal dynamics of complex events.

The second challenge is generalizing to unseen anomalies, since abnormal events are often absent from the training data.

The third challenge is efficiently modeling long video sequences, which remains computationally demanding for many existing approaches.

These challenges become even more difficult in practice because anomalies are naturally diverse, ambiguous, and rare.

As a result, developing robust temporal and semantic representations becomes essential for reliable anomaly detection.

An additional challenge comes from the weakly supervised setting.

In WSVAD, only video-level labels are available during training, making the task even more challenging because the model must infer which temporal segments are responsible for the anomalous event.

===

These challenges naturally lead to the following question: how can we learn temporal representations that are both effective and efficient while maintaining strong generalization? This is the research gap addressed in this work.

==

So we know that transformer and VLD Models have progress, however they are expensive for onger videos. On the other side, SSMs such as Mamba arise from it's effiency, however, the implementation on VAD is newer and there is still much are for improvement, as Mamba cannot be used direftly for VAD. Finally, Generaliation requieres additional context in order to improve.

==


As we have seen in recent years, Transformers have represented the state of the art, and when combined with Vision-Language Models, they have achieved significant progress in Video Anomaly Detection. However, they remain computationally expensive, especially for long video sequences.

On the other hand, State Space Models, particularly Mamba, have emerged as a promising alternative due to their computational efficiency. However, their application to Video Anomaly Detection is still in its early stages and requires task-specific adaptations.

Finally, improving generalization to unseen anomalies remains an open challenge, as recognizing these events often requires richer temporal representations and semantic information.

With this in mind, the starting point of this research is not a single architecture, but rather understanding which temporal properties an effective representation should learn and which architecture can best capture them.

So, the objective of this research is to investigate how robust temporal representations can improve anomaly detection in long and complex videos.

==

So the central hypothesis is that:

If temporal representations effectively integrate complementary semantic or contextual information, they can better distinguish normal from anomalous patterns, improving the detection of weak, ambiguous, or previously unseen anomalies.

If this hypothesis is correct, it should lead to more efficient, semantic-aware, and robust Video Anomaly Detection systems.

=

To investigate this hypothesis, I propose a three-stage research methodology.

The first stage focuses on understanding the current state of the art in Video Anomaly Detection. During this stage, I will analyze CNN-based, Transformer-based, and Vision-Language approaches, identifying their strengths, limitations, and temporal modeling strategies.

The second stage investigates State Space Models, particularly Mamba and its variants, as efficient alternatives for temporal modeling. This stage aims to understand what temporal properties these models capture and under which conditions they provide advantages over attention-based methods.

The final stage explores the integration of semantic information from Vision-Language Models with efficient temporal representations. The objective is to evaluate whether complementary semantic information improves the recognition of weak, ambiguous, and previously unseen anomalies.

==

Finally, the proposed approach will be evaluated on three widely used benchmark datasets: UCF-Crime, XD-Violence, and ShanghaiTech.

The evaluation will be conducted using frame-level AUC and Average Precision to measure detection performance.

In addition, computational efficiency and robustness will be evaluated to determine whether the proposed method provides a better balance between accuracy, efficiency, and generalization.

The results will be compared with existing state-of-the-art methods to assess the effectiveness of the proposed approach.

==

I expect this research to contribute from both scientific and practical perspectives.

From a scientific perspective, I expect to provide a better understanding of temporal representation learning for Weakly Supervised Video Anomaly Detection, as well as new insights into the integration of efficient temporal modeling and semantic information.

From a practical perspective, I expect this research to contribute to the development of more robust, scalable, and generalizable video anomaly detection systems under weak supervision.

==


Thank you very much for your attention. I will be happy to answer any questions.

